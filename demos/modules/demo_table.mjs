/**
 * NAME: demo_tables.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.12.0
 * BLD.: 20230207
 */

import {
	TABLE_NAMES_F,
	DEMO_TITLE_OPTS,
	DEMO_TITLE_TEXT,
	DEMO_TITLE_TEXTBK,
	BASE_OPTS_SUBTITLE,
	BASE_TABLE_OPTS,
	BASE_TEXT_OPTS_L,
	BASE_TEXT_OPTS_R,
	LOREM_IPSUM,
} from "./enums.mjs";

export function genSlides_Table(pptx) {
	pptx.addSection({ title: "Tables" });
	genSlide01(pptx);
	genSlide02(pptx);
	genSlide03(pptx);
	genSlide04(pptx);
	genSlide05(pptx);
	genSlide06(pptx);

	pptx.addSection({ title: "Tables: Auto-Paging" });
	genSlide07(pptx);

	pptx.addSection({ title: "Tables: Auto-Paging Complex" });
	genSlide08(pptx);

	pptx.addSection({ title: "Tables: Auto-Paging Calc" });
	genSlide09(pptx);

	pptx.addSection({ title: "Tables: QA" });
	genSlide10(pptx);
}

/**
 * SLIDE 1: Table text alignment and cell styles
 * @param {PptxGenJS} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Tables" });

	slide.addTable([[{ text: "Table Examples 1", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs:\nhttps://gitbrent.github.io/PptxGenJS/docs/api-tables.html");

	// DEMO: align/valign
	let objOpts1 = { x: 0.5, y: 0.7, w: 4, h: 0.3, margin: 0, fontSize: 18, fontFace: "Arial", color: "0088CC" };
	slide.addText("Cell Text Alignment:", objOpts1);

	let arrTabRows1 = [
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
		border: { pt: "1", color: "BBCCDD" },
	});
	// Pass default cell style as tabOpts, then just style/override individual cells as needed

	// DEMO: cell styles
	let objOpts2 = { x: 6.0, y: 0.7, w: 4, h: 0.3, margin: 0, fontSize: 18, fontFace: "Arial", color: "0088CC" };
	slide.addText("Cell Styles:", objOpts2);

	let arrTabRows2 = [
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
			{ text: "Underline", options: { fill: { color: "336699" }, underline: { style: "sng" } } },
			{ text: "0.15 margin", options: { fill: { color: "6699CC" }, margin: 0.15 } },
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
		valign: "center",
		align: "center",
		border: { pt: "1", color: "FFFFFF" },
	});

	// DEMO: Row/Col Width/Heights
	let objOpts3 = { x: 0.5, y: 3.6, h: 0.3, margin: 0, fontSize: 18, fontFace: "Arial", color: "0088CC" };
	slide.addText("Row/Col Heights/Widths:", objOpts3);

	let arrTabRows3 = [
		[{ text: "1x1" }, { text: "2x1" }, { text: "2.5x1" }, { text: "3x1" }, { text: "4x1" }],
		[{ text: "1x2" }, { text: "2x2" }, { text: "2.5x2" }, { text: "3x2" }, { text: "4x2" }],
	];
	slide.addTable(arrTabRows3, {
		x: 0.5,
		y: 4.0,
		rowH: [1, 2],
		colW: [1, 2, 2.5, 3, 4],
		fill: { color: "F7F7F7" },
		color: "6c6c6c",
		fontSize: 14,
		valign: "center",
		align: "center",
		border: { pt: "1", color: "BBCCDD" },
	});
}

/**
 * SLIDE 2: Table row/col-spans
 * @param {PptxGenJS} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Tables" });

	slide.addTable([[{ text: "Table Examples 2", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], { x: "4%", y: "2%", w: "95%", h: "4%" }); // QA: this table's x,y,w,h all using %
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");

	// DEMO: Rowspans/Colspans
	let optsSub = JSON.parse(JSON.stringify(BASE_OPTS_SUBTITLE));
	slide.addText("Colspans/Rowspans:", optsSub);

	let tabOpts1 = {
		x: 0.67,
		y: 1.1,
		w: "90%",
		h: 2,
		fill: { color: "F9F9F9" },
		color: "3D3D3D",
		fontSize: 16,
		border: { pt: 4, color: "FFFFFF" },
		align: "center",
		valign: "middle",
	};
	let arrTabRows1 = [
		[
			{ text: "A1\nA2", options: { rowspan: 2, fill: { color: "99FFCC" } } },
			{ text: "B1" },
			{ text: "C1 -> D1", options: { colspan: 2, fill: { color: "99FFCC" } } },
			{ text: "E1" },
			{ text: "F1\nF2\nF3", options: { rowspan: 3, fill: { color: "99FFCC" } } },
		],
		["B2", "C2", "D2", "E2"],
		["A3", "B3", "C3", "D3", "E3"],
	];
	// NOTE: Follow HTML conventions for colspan/rowspan cells - cells spanned are left out of arrays - see above
	// The table above has 6 columns, but each of the 3 rows has 4-5 elements as colspan/rowspan replacing the missing ones
	// (e.g.: there are 5 elements in the first row, and 6 in the second)
	slide.addTable(arrTabRows1, tabOpts1);

	let tabOpts2 = {
		x: 0.5,
		y: 3.3,
		w: 12.4,
		h: 1.5,
		fontSize: 14,
		fontFace: "Courier",
		align: "center",
		valign: "middle",
		fill: { color: "F9F9F9" },
		border: { pt: "1", color: "c7c7c7" },
	};
	let arrTabRows2 = [
		[
			{ text: "A1\n--\nA2", options: { rowspan: 2, fill: { color: "99FFCC" } } },
			{ text: "B1\n--\nB2", options: { rowspan: 2, fill: { color: "99FFCC" } } },
			{ text: "C1 -> D1", options: { colspan: 2, fill: { color: "9999FF" } } },
			{ text: "E1 -> F1", options: { colspan: 2, fill: { color: "9999FF" } } },
			"G1",
		],
		["C2", "D2", "E2", "F2", "G2"],
	];
	slide.addTable(arrTabRows2, tabOpts2);

	// BTM-LEFT
	let tabOpts3 = {
		x: 0.5,
		y: 5.15,
		w: 6.0,
		h: 2,
		margin: 0.05,
		align: "center",
		valign: "middle",
		fontSize: 16,
		border: { pt: "2", color: pptx.colors.TEXT2 },
		fill: { color: "F1F1F1" },
	};
	let arrTabRows3 = [
		[
			{ text: "A1\nA2\nA3", options: { rowspan: 3, fill: { color: pptx.colors.ACCENT6 } } },
			{ text: "B1\nB2", options: { rowspan: 2, fill: { color: pptx.colors.ACCENT2 } } },
			{ text: "C1", options: { fill: { color: pptx.colors.ACCENT4 } } },
		],
		[{ text: "C2", options: { fill: { color: pptx.colors.ACCENT4 } } }],
		[{ text: "B3 -> C3", options: { colspan: 2, fill: { color: pptx.colors.ACCENT5 } } }],
	];
	slide.addTable(arrTabRows3, tabOpts3);

	// BTM-RIGHT
	let tabOpts4 = {
		x: 6.93,
		y: 5.15,
		w: 6.0,
		h: 2,
		margin: 0,
		align: "center",
		valign: "middle",
		fontSize: 16,
		border: { pt: "1", color: pptx.colors.TEXT2 },
		fill: { color: "f2f9fc" },
	};
	let arrTabRows4 = [
		[
			{ text: "A1", options: { fill: { color: pptx.colors.ACCENT4, transparency: 25 } } },
			{ text: "B1\nB2", options: { rowspan: 2, fill: { color: pptx.colors.ACCENT2, transparency: 25 } } },
			{ text: "C1\nC2\nC3", options: { rowspan: 3, fill: { color: pptx.colors.ACCENT6, transparency: 25 } } },
		],
		[{ text: "A2", options: { fill: { color: pptx.colors.ACCENT4, transparency: 25 } } }],
		[{ text: "A3 -> B3", options: { colspan: 2, fill: { color: pptx.colors.ACCENT5, transparency: 25 } } }],
	];
	slide.addTable(arrTabRows4, tabOpts4);
}

/**
 * SLIDE 3: Super rowspan/colspan demo
 * @param {PptxGenJS} pptx
 */
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Tables" });

	slide.addTable([[{ text: "Table Examples 3", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");

	// DEMO: Rowspans/Colspans ----------------------------------------------------------------
	let optsSub = JSON.parse(JSON.stringify(BASE_OPTS_SUBTITLE));
	slide.addText("Extreme Colspans/Rowspans:", optsSub);

	let optsRowspan2 = { rowspan: 2, fill: { color: "99FFCC" } };
	let optsRowspan3 = { rowspan: 3, fill: { color: "99FFCC" } };
	let optsRowspan4 = { rowspan: 4, fill: { color: "99FFCC" } };
	let optsRowspan5 = { rowspan: 5, fill: { color: "99FFCC" } };
	let optsColspan2 = { colspan: 2, fill: { color: "9999FF" } };
	let optsColspan3 = { colspan: 3, fill: { color: "9999FF" } };
	let optsColspan4 = { colspan: 4, fill: { color: "9999FF" } };
	let optsColspan5 = { colspan: 5, fill: { color: "9999FF" } };

	let arrTabRows5 = [
		[
			"A1",
			"B1",
			"C1",
			"D1",
			"E1",
			"F1",
			"G1",
			"H1",
			{ text: "I1\n-\nI2\n-\nI3\n-\nI4\n-\nI5", options: optsRowspan5 },
			{ text: "J1 -> K1 -> L1 -> M1 -> N1", options: optsColspan5 },
		],
		[
			{ text: "A2\n--\nA3", options: optsRowspan2 },
			{ text: "B2 -> C2 -> D2", options: optsColspan3 },
			{ text: "E2 -> F2", options: optsColspan2 },
			{ text: "G2\n-\nG3\n-\nG4", options: optsRowspan3 },
			"H2",
			"J2",
			"K2",
			"L2",
			"M2",
			"N2",
		],
		[{ text: "B3\n-\nB4\n-\nB5", options: optsRowspan3 }, "C3", "D3", "E3", "F3", "H3", "J3", "K3", "L3", "M3", "N3"],
		[
			{ text: "A4\n--\nA5", options: optsRowspan2 },
			{ text: "C4 -> D4 -> E4 -> F4", options: optsColspan4 },
			"H4",
			{ text: "J4 -> K4 -> L4", options: optsColspan3 },
			{ text: "M4\n--\nM5", options: optsRowspan2 },
			{ text: "N4\n--\nN5", options: optsRowspan2 },
		],
		["C5", "D5", "E5", "F5", { text: "G5 -> H5", options: { colspan: 2, fill: { color: "9999FF" } } }, "J5", "K5", "L5"],
	];

	let taboptions5 = { x: 0.6, y: 1.3, w: "90%", h: 5.5, margin: 0, fontSize: 14, align: "center", valign: "middle", border: { pt: "1" } };

	slide.addTable(arrTabRows5, taboptions5);
}

/**
 * SLIDE 4: Cell Formatting / Cell Margins
 * @param {PptxGenJS} pptx
 */
function genSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Tables" });

	slide.addTable([[{ text: "Table Examples 4", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");

	// Cell Margins
	let optsSub = JSON.parse(JSON.stringify(BASE_OPTS_SUBTITLE));
	slide.addText("Cell Margins:", optsSub);

	slide.addTable([["margin:0"]], { x: 0.5, y: 1.1, margin: 0, w: 1.2, fill: "FFFCCC", border: { pt: 0 } });
	slide.addTable([["margin:[0, 0, 0, 0.3]"]], { x: 2.5, y: 1.1, margin: [0, 0, 0, 0.3], w: 2.0, fill: "FFFCCC", align: "right" });
	slide.addTable([["margin:0.05"]], { x: 5.5, y: 1.1, margin: 0.05, w: 1.0, fill: pptx.SchemeColor.background2 });
	slide.addTable([["margin:[0.6, 0.05, 0.05, 0.3]"]], { x: 7.1, y: 1.1, margin: [0.6, 0.05, 0.05, 0.3], w: 2.6, fill: "F1F1F1" });
	slide.addTable([["margin:[0.45, 0.05, 0.05, 0.45]"]], { x: 10.1, y: 1.1, margin: [0.45, 0.05, 0.05, 0.45], w: 2.6, fill: "F1F1F1" });

	slide.addTable(
		[
			[
				{ text: "no border and number zero", options: { margin: 0.05 } },
				{ text: 0, options: { margin: 0.05 } },
			],
		],
		{ x: 0.5, y: 1.9, fill: { color: "f2f9fc" }, border: { type: "none" }, colW: [2.5, 0.5] }
	);
	slide.addTable([[{ text: "text-obj margin:0", options: { margin: 0 } }]], { x: 4.0, y: 1.9, w: 2, fill: { color: "f2f9fc" } });

	// Test margin option when using both plain and text object cells
	let arrTextObjects = [
		["Plain text", "Cell 2", 3],
		[
			{ text: "Text Objects", options: { color: "99ABCC", align: "right" } },
			{ text: "2nd cell", options: { color: "0000EE", align: "center" } },
			{ text: "Hyperlink", options: { hyperlink: { url: "https://github.com/gitbrent/pptxgenjs" } } },
		],
	];
	slide.addTable(arrTextObjects, { x: 0.5, y: 2.7, w: 12.25, margin: 8, fill: { color: "F1F1F1" }, border: { pt: 1, color: "696969" } }); // DEPRECATED: `margin` in points

	// Complex/Compound border
	optsSub.y = 3.9;
	slide.addText("Complex Cell Borders:", optsSub);
	let arrBorder1 = [
		{ color: "FF0000", pt: 1 },
		{ color: "00ff00", pt: 3 },
		{ color: "0000ff", pt: 5 },
		{ color: "9e9e9e", pt: 7 },
	];
	slide.addTable([["Borders 4!"]], {
		x: 0.5,
		y: 4.3,
		w: 6,
		rowH: 1.5,
		fill: pptx.SchemeColor.background2,
		color: "3D3D3D",
		fontSize: 18,
		border: arrBorder1,
		align: "center",
		valign: "middle",
	});
	let arrBorder2 = [{ type: "dash", color: "ff0000", pt: 2 }, null, { type: "dash", color: "0000ff", pt: 5 }, null];
	slide.addTable([["Borders 2!"]], {
		x: 6.75,
		y: 4.3,
		w: 6,
		rowH: 1.5,
		fill: pptx.SchemeColor.background2,
		color: "3D3D3D",
		fontSize: 18,
		border: arrBorder2,
		align: "center",
		valign: "middle",
	});

	// Special char check
	optsSub.y = 6.1;
	slide.addText("Escaped Special Chars:", optsSub);
	let arrTabRows3 = [["<", ">", '"', "'", "&", "plain"]];
	slide.addTable(arrTabRows3, {
		x: 0.5,
		y: 6.5,
		w: 12.3,
		rowH: 0.5,
		fill: { color: "F9F9F9" },
		color: "3D3D3D",
		border: { pt: 1, color: "FFFFFF" },
		align: "center",
		valign: "middle",
	});
}

/**
 * SLIDE 5: Cell Word-Level Formatting
 * @param {PptxGenJS} pptx
 */
function genSlide05(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Tables" });

	slide.addTable([[{ text: "Table Examples 5", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");

	slide.addText(
		"The following textbox and table cell use the same array of text/options objects, making word-level formatting familiar and consistent across the library.",
		{ x: 0.5, y: 0.5, w: "95%", h: 0.5, margin: 0.1, fontSize: 14 }
	);
	slide.addText(
		"[\n" +
			"  { text:'1st line', options:{ fontSize:24, color:'99ABCC', align:'right',  breakLine:true } },\n" +
			"  { text:'2nd line', options:{ fontSize:36, color:'FFFF00', align:'center', breakLine:true } },\n" +
			"  { text:'3rd line', options:{ fontSize:48, color:'0088CC', align:'left'    } }\n" +
			"]",
		{ x: 1, y: 1.1, w: 11, h: 1.25, margin: 0.1, fontFace: "Courier", fontSize: 13, fill: { color: "F1F1F1" }, color: "333333" }
	);

	// Textbox: Text word-level formatting
	slide.addText("Textbox:", { x: 1, y: 2.8, w: 3, h: 2, fontSize: 18, fontFace: "Arial", color: "0088CC" });

	let arrTextObjects = [
		{ text: "1st line", options: { fontSize: 24, color: "99ABCC", align: "right", breakLine: true } },
		{ text: "2nd line", options: { fontSize: 36, color: "FFFF00", align: "center", breakLine: true } },
		{ text: "3rd line", options: { fontSize: 48, color: "0088CC", align: "left" } },
	];
	slide.addText(arrTextObjects, { x: 2.5, y: 2.8, w: 9.5, h: 2, margin: 0.1, fill: { color: "232323" } });

	// Table cell: Use the exact same code from addText to do the same word-level formatting within a cell
	slide.addText("Table:", { x: 1, y: 5, w: 3, h: 2, fontSize: 18, fontFace: "Arial", color: "0088CC" });

	let opts2 = { x: 2.5, y: 5, h: 2, align: "center", valign: "middle", colW: [1.5, 1.5, 6.5], border: { pt: "1" }, fill: { color: "F1F1F1" } };
	let arrTabRows = [
		[
			{ text: "Cell 1A", options: { fontFace: "Arial" } },
			{ text: "Cell 1B", options: { fontFace: "Courier" } },
			{ text: arrTextObjects, options: { fill: { color: "232323" } } },
		],
	];
	slide.addTable(arrTabRows, opts2);
}

/**
 * SLIDE 6: Cell Word-Level Formatting
 * @param {PptxGenJS} pptx
 */
function genSlide06(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Tables" });

	slide.addTable([[{ text: "Table Examples 6", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");

	let optsSub = JSON.parse(JSON.stringify(BASE_OPTS_SUBTITLE));
	slide.addText("Table Cell Word-Level Formatting:", optsSub);

	// EX 1:
	let arrCell1 = [{ text: "Cell\n#1", options: { color: "0088cc" } }];
	let arrCell2 = [
		{ text: "Red ", options: { color: "FF0000" } },
		{ text: "Green ", options: { color: "00FF00" } },
		{ text: "Blue", options: { color: "0000FF" } },
	];
	let arrCell3 = [{ text: "google", options: { bullet: true, color: pptx.colors.ACCENT1, hyperlink: { url: "https://www.google.com" } } }];

	let arrCell4 = [{ text: "Numbers\nNumbers\nNumbers", options: { color: "0088cc", bullet: { type: "number" } } }];
	slide.addTable(
		[
			[
				{ text: arrCell1 },
				{ text: arrCell2, options: { valign: "middle" } },
				{ text: arrCell3, options: { valign: "middle" } },
				{ text: arrCell4, options: { valign: "bottom" } },
			],
		],
		{ x: 0.6, y: 1.25, w: 12, h: 3, fontSize: 24, border: { pt: "1" }, fill: { color: "F1F1F1" } }
	);

	// EX 2:
	slide.addTable(
		[
			[
				{
					text: [
						{ text: "I am a text object with bullets ", options: { color: "CC0000", bullet: { code: "2605" } } },
						{ text: "and i am the next text object", options: { color: "00CD00", bullet: { code: "25BA" } } },
						{ text: "Final text object w/ bullet:true", options: { color: "0000AB", bullet: true } },
					],
				},
				{
					text: [
						{ text: "Cell", options: { fontSize: 36, align: "left", color: "8648cd" } },
						{ text: "#2", options: { fontSize: 60, align: "right", color: "CD0101" } },
					],
				},
				{
					text: [
						{ text: "Cell", options: { fontSize: 36, fontFace: "Courier", color: "dd0000", breakLine: true } },
						{ text: "#", options: { fontSize: 60, color: "8648cd" } },
						{ text: "3", options: { fontSize: 60, fontFace: "Times", color: "33ccef" } },
					],
				},
			],
		],
		{ x: 0.6, y: 4.75, h: 2, fontSize: 24, colW: [8, 2, 2], valign: "middle", border: { pt: "1" }, fill: { color: "F1F1F1" } }
	);
}

/**
 * SLIDE 7[...]: Table auto-paging
 * @param {PptxGenJS} pptx
 */
function genSlide07(pptx) {
	let slide = null;

	// STEP 1: Build data
	let arrRows = [];
	let arrText = [];
	let arrRowsHead1 = [];
	let arrRowsHead2 = [[{ text: "Title Header", options: { fill: "0088cc", color: "ffffff", align: "center", bold: true, colspan: 3, colW: 4 } }]];
	{
		arrRows.push([
			{ text: "ID#", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
			{ text: "First Name", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
			{ text: "Lorum Ipsum", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
		]);
		TABLE_NAMES_F.forEach((name, idx) => {
			let strText = idx == 0 ? LOREM_IPSUM.substring(0, 100) : LOREM_IPSUM.substring(idx * 100, idx * 200);
			arrRows.push([idx + 1, name, strText]);
			arrText.push([strText]);
		});
		arrRows.forEach((row, idx) => {
			if (idx < 6) arrRowsHead1.push(row);
		});
		arrRows.forEach((row, idx) => {
			if (idx < 6) arrRowsHead2.push(row);
		});
	}

	// EX-1: "Basic Auto-Paging Example"
	{
		slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging" });
		slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");
		slide.addText(
			[
				{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
				{ text: "Basic Auto-Paging Example", options: DEMO_TITLE_OPTS },
			],
			{ x: 0.5, y: 0.13, w: "90%" }
		);
		slide.addTable(arrRows, { x: 0.5, y: 0.5, colW: [0.75, 1.75, 10], margin: 0.05, border: { color: "CFCFCF" }, autoPage: true });
	}

	// EX-2: "Paging with smaller table area (50% width)"
	{
		slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging" });
		slide.addText(
			[
				{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
				{ text: "Paging with smaller table area (50% width)", options: DEMO_TITLE_OPTS },
			],
			{ x: 0.5, y: 0.13, w: "90%" }
		);
		slide.addTable(arrRows, {
			x: "50%",
			y: 0.5,
			colW: [0.75, 1.75, 4],
			margin: 0.05,
			border: { color: "CFCFCF" },
			autoPage: true,
			verbose: false,
		});
	}

	// EX-3: "Master Page with Auto-Paging"
	{
		slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging", masterName: "MASTER_AUTO_PAGE_TABLE_PLACEHOLDER" });
		slide.addText(
			[
				{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
				{ text: "Master Page with Auto-Paging", options: DEMO_TITLE_OPTS },
			],
			{ x: 0.5, y: 0.13, w: "90%" }
		);
		slide.addText("Auto-Paging table", { placeholder: "footer" });
		slide.addTable(arrRows, { x: 1.0, y: 0.6, colW: [0.75, 1.75, 7], margin: 0.05, border: { color: "CFCFCF" }, autoPage: true });
		// HOWTO: In cases where you want to add custom text, placeholders, etc. to slidemasters, a reference to these slide(s) is needed
		// HOWTO: Use the `newAutoPagedSlides` to access references (see [Issue #625](https://github.com/gitbrent/PptxGenJS/issues/625))
		slide.newAutoPagedSlides.forEach((slide) => slide.addText("Auto-Paging table continued...", { placeholder: "footer" }));
	}

	// EX-4: "Auto-Paging Disabled"
	{
		slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging" });
		slide.addText(
			[
				{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
				{ text: "Auto-Paging Disabled", options: DEMO_TITLE_OPTS },
			],
			{ x: 0.5, y: 0.13, w: "90%" }
		);
		slide.addTable(arrRows, { x: 1.0, y: 0.6, colW: [0.75, 1.75, 7], margin: 0.05, border: { color: "CFCFCF" } }); // Negative-Test: no `autoPage:false`
	}

	// EX-5: "Start at `{ y: 4.0 }`, subsequent slides start at slide top margin"
	{
		slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging", masterName: "MARGIN_SLIDE" });
		slide.addText(
			[
				{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
				{ text: "Start at `{ y:4.0 }`, subsequent slides start at slide top margin", options: DEMO_TITLE_OPTS },
			],
			{ x: 3.0, y: 0.75, w: "75%", h: 0.5 }
		);
		slide.addTable(arrRows, {
			x: 3.0,
			y: 4.0,
			colW: [0.75, 1.75, 7],
			margin: 0.05,
			border: { color: "CFCFCF" },
			fontFace: "Arial",
			autoPage: true,
			verbose: false,
		});
	}

	// EX-6: "Start at `{ y: 4.0 }`, subsequent slides start at `{ autoPageSlideStartY: 1.5 }`"
	{
		slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging", masterName: "MARGIN_SLIDE_STARTY15" });
		slide.addText(
			[
				{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
				{ text: "Start at `{ y: 4.0 }`, subsequent slides start at `{ autoPageSlideStartY: 1.5 }`", options: DEMO_TITLE_OPTS },
			],
			{ x: 3.0, y: 0.75, w: "75%", h: 0.5 }
		);
		slide.addTable(arrRows, {
			x: 3.0,
			y: 4.0,
			colW: [0.75, 1.75, 7],
			margin: 0.05,
			border: { color: "CFCFCF" },
			autoPage: true,
			autoPageSlideStartY: 1.5,
			autoPageCharWeight: 0.15,
			verbose: false,
		});
	}

	// EX-7: `autoPageRepeatHeader` option demos
	{
		pptx.addSection({ title: "Tables: Auto-Paging Repeat Header" });
		slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging Repeat Header" });
		slide.addText(
			[
				{ text: "Table Examples: `autoPageHeaderRows`", options: DEMO_TITLE_TEXTBK },
				{ text: "no `autoPageHeaderRows` prop", options: DEMO_TITLE_OPTS },
			],
			{ x: 0.23, y: 0.13, w: 4, h: 0.4 }
		);
		slide.addTable(arrRowsHead1, {
			x: 0.23,
			y: 0.6,
			colW: [0.5, 1.0, 2.5],
			margin: 0.05,
			border: { color: "CFCFCF" },
			autoPage: true,
			autoPageRepeatHeader: true,
			autoPageSlideStartY: 0.6,
		});

		slide.addText(
			[
				{ text: "Table Examples: autoPageHeaderRows", options: DEMO_TITLE_TEXTBK },
				{ text: "`{ autoPageHeaderRows: 1 }`", options: DEMO_TITLE_OPTS },
			],
			{ x: 4.75, y: 0.13, w: 4, h: 0.4 }
		);
		slide.addTable(arrRowsHead1, {
			x: 4.75,
			y: 0.6,
			colW: [0.5, 1.0, 2.5],
			margin: 0.05,
			border: { color: "CFCFCF" },
			autoPage: true,
			autoPageRepeatHeader: true,
			autoPageHeaderRows: 1,
			autoPageSlideStartY: 0.6,
		});

		slide.addText(
			[
				{ text: "Table Examples: autoPageHeaderRows", options: DEMO_TITLE_TEXTBK },
				{ text: "`{ autoPageHeaderRows: 2 }`", options: DEMO_TITLE_OPTS },
			],
			{ x: 9.1, y: 0.13, w: 4, h: 0.4 }
		);
		slide.addTable(arrRowsHead2, {
			x: 9.1,
			y: 0.6,
			colW: [0.5, 1.0, 2.5],
			margin: 0.05,
			border: { color: "CFCFCF" },
			autoPage: true,
			autoPageRepeatHeader: true,
			autoPageHeaderRows: 2,
			autoPageSlideStartY: 0.6,
		});
	}

	// EX-8: `autoPageLineWeight` option demos
	{
		pptx.addSection({ title: "Tables: Auto-Paging LineWeight" });
		slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging LineWeight" });
		slide.addText(
			[
				{ text: "Table Examples: Line Weight Options", options: DEMO_TITLE_TEXTBK },
				{ text: "autoPageLineWeight:0.0", options: DEMO_TITLE_OPTS },
			],
			{ x: 0.23, y: 0.13, w: 4, h: 0.4 }
		);
		slide.addTable(arrText, { x: 0.23, y: 0.6, w: 4, margin: 0.05, border: { color: "CFCFCF" }, autoPage: true, autoPageLineWeight: 0.0 });

		slide.addText(
			[
				{ text: "Table Examples: Line Weight Options", options: DEMO_TITLE_TEXTBK },
				{ text: "autoPageLineWeight:0.5", options: DEMO_TITLE_OPTS },
			],
			{ x: 4.75, y: 0.13, w: 4, h: 0.4 }
		);
		slide.addTable(arrText, { x: 4.75, y: 0.6, w: 4, margin: 0.05, border: { color: "CFCFCF" }, autoPage: true, autoPageLineWeight: 0.5 });

		slide.addText(
			[
				{ text: "Table Examples: Line Weight Options", options: DEMO_TITLE_TEXTBK },
				{ text: "autoPageLineWeight:-0.5", options: DEMO_TITLE_OPTS },
			],
			{ x: 9.1, y: 0.13, w: 4, h: 0.4 }
		);
		slide.addTable(arrText, { x: 9.1, y: 0.6, w: 4, margin: 0.05, border: { color: "CFCFCF" }, autoPage: true, autoPageLineWeight: -0.5 });
	}

	// EX-9: `autoPageCharWeight` option demos
	{
		pptx.addSection({ title: "Tables: Auto-Paging CharWeight" });
		slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging CharWeight" });
		slide.addText(
			[
				{ text: "Table Examples: Char Weight Options", options: DEMO_TITLE_TEXTBK },
				{ text: "autoPageCharWeight:0.0", options: DEMO_TITLE_OPTS },
			],
			{ x: 0.23, y: 0.13, w: 4, h: 0.4 }
		);
		slide.addTable(arrText, { x: 0.23, y: 0.6, w: 4, margin: 0.05, border: { color: "CFCFCF" }, autoPage: true, autoPageCharWeight: 0.0 });

		slide.addText(
			[
				{ text: "Table Examples: Char Weight Options", options: DEMO_TITLE_TEXTBK },
				{ text: "autoPageCharWeight:0.25", options: DEMO_TITLE_OPTS },
			],
			{ x: 4.75, y: 0.13, w: 4, h: 0.4 }
		);
		slide.addTable(arrText, { x: 4.75, y: 0.6, w: 4, margin: 0.05, border: { color: "CFCFCF" }, autoPage: true, autoPageCharWeight: 0.25 });

		slide.addText(
			[
				{ text: "Table Examples: Char Weight Options", options: DEMO_TITLE_TEXTBK },
				{ text: "autoPageCharWeight:-0.25", options: DEMO_TITLE_OPTS },
			],
			{ x: 9.1, y: 0.13, w: 4, h: 0.4 }
		);
		slide.addTable(arrText, { x: 9.1, y: 0.6, w: 4, margin: 0.05, border: { color: "CFCFCF" }, autoPage: true, autoPageCharWeight: -0.25 });
	}
}

/**
 * SLIDE 8[...]: Table auto-paging with complex text array (unsupported until 3.7.2/3.8.0)
 * @param {PptxGenJS} pptx
 * @since 3.8.0
 */
function genSlide08(pptx) {
	let slide = null;
	let arrRows = [];

	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging Complex" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");
	slide.addText([{ text: "Table Examples: Auto-Paging Using Complex Text Example", options: DEMO_TITLE_TEXTBK }], {
		x: 0.23,
		y: 0.13,
		w: 8,
		h: 0.4,
	});

	// ------------

	arrRows.push([
		{ text: "ID#", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
		{ text: "First Name", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
		{ text: "Lorum Ipsum", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
	]);
	TABLE_NAMES_F.forEach((name, idx) => {
		let strText = idx == 0 ? LOREM_IPSUM.substring(0, 100) : LOREM_IPSUM.substring(idx * 100, idx * 200);
		arrRows.push([
			{ text: idx.toString(), options: { align: "center" } },
			{ text: name },
			{ text: [{ text: "Title", options: { bold: true, color: "FF0000", breakLine: true } }, { text: `>${strText}<` }] },
		]);
	});
	slide.addTable(arrRows, {
		x: 0.5,
		y: 0.5,
		w: 8,
		colW: [1, 1, 6],
		border: { color: "CFCFCF" },
		autoPage: true,
		autoPageRepeatHeader: true,
		verbose: false,
	});
}

/**
 * SLIDE 9[...]: Tightly calculated/labels rows and cells for precision auto-paging dev & test
 * @param {PptxGenJS} pptx
 * @since 3.8.0
 */
function genSlide09(pptx) {
	let slide = null;
	let arrRows = [];

	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging Calc" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");
	slide.addText([{ text: "Table Examples: Auto-Paging Calculations", options: DEMO_TITLE_TEXTBK }], {
		x: 0.23,
		y: 0.13,
		w: 8,
		h: 0.4,
	});

	// ------------

	arrRows.push([
		{ text: "TH1", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
		{ text: "TH2", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
		{ text: "TH3", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
	]);
	for (let rowIdx = 0; rowIdx < 9; rowIdx++) {
		let col3Lines = [{ text: "Complex-Title", options: { bold: true, color: "FF0000", breakLine: true } }];
		for (let lineIdx = 0; lineIdx < 9; lineIdx++) {
			col3Lines.push({ text: `This is ROW#:${rowIdx + 1} LNE#:${lineIdx + 1}`, options: { breakLine: true } });
		}
		arrRows.push([{ text: "" }, { text: "" }, { text: col3Lines }]);
	}
	slide.addTable(arrRows, {
		x: 0.5,
		y: 0.75,
		w: 8,
		colW: [1, 1, 6],
		border: { color: "CFCFCF" },
		autoPage: true,
		autoPageRepeatHeader: true,
		verbose: false,
	});
}

/**
 * SLIDE 10[...]: Test paging with a single row
 * @param {PptxGenJS} pptx
 * @since 3.9.0
 */
function genSlide10(pptx) {
	let slide = null;

	// SLIDE 1:
	{
		slide = pptx.addSlide({ sectionTitle: "Tables: QA" });

		let projRows = [
			[
				{ text: "id", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "First item Desc", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Impact", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Owner", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Created Date", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Due Date", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Status", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Update", options: { bold: true, fill: "1F3864", color: "ffffff" } },
			],
			[
				{ text: "1" },
				{
					text: "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, arcu. In enim justo, rhoncus ut, imperdiet a, venenatis vitae, justo. Nullam dictum felis eu pede mollis pretium. Integer tincidunt. Cras dapibus. Vivamus elementum semper nisi. Aenean vulputate eleifend tellus. Aenean leo ligula, porttitor eu, consequat vitae, eleifend ac, enim. Aliquam lorem ante, dapibus in, viverra quis, feugiat a, tellus. Phasellus viverra nulla ut metus varius laoreet. Quisque rutrum. Aenean imperdiet. Etiam ultricies nisi vel augue. Curabitur ullamcorper ultricies nisi. Nam eget dui. Etiam rhoncus. Maecenas tempus, tellus eget condimentum rhoncus, sem quam semper libero, sit amet adipiscing sem neque sed ipsum. Nam quam nunc, blandit vel, luctus pulvinar, hendrerit id, lorem. Maecenas nec odio et ante tincidunt tempus. Donec vitae sapien ut libero venenatis faucibus. Nullam quis ante. Etiam sit amet orci eget eros faucibus tincidunt. Duis leo. Sed fringilla mauris sit amet nibh. Donec sodales sagittis magna.",
				},
				{ text: "Adam" },
				{ text: "Aenean commodo ligula eget dolor. Aenean massa." },
				{ text: "20-10-2021" },
				{ text: "01-11-2021" },
				{ text: "Pending" },
				{ text: "24-10-2021" },
			],
		];

		slide.addTable(projRows, {
			x: 0.4,
			y: 5.25,
			colW: [0.5, 1.8, 5, 0.9, 1.0, 0.95, 0.8, 1.5],
			border: { pt: 0.1, color: "818181" },
			align: "left",
			valign: "middle",
			fontFace: "Segoe UI",
			fontSize: 8,
			autoPage: true,
			autoPageRepeatHeader: true,
			autoPageLineWeight: -0.4,
			verbose: true,
		});
	}

	// SLIDE 2:
	{
		slide = pptx.addSlide({ sectionTitle: "Tables: QA" });

		let projRows2 = [
			[
				{ text: "id", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "First item Desc", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Impact", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Owner", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Created Date", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Due Date", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Status", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Update", options: { bold: true, fill: "1F3864", color: "ffffff" } },
			],
			[
				{ text: "1" },
				{
					text: "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, arcu. In enim justo, rhoncus ut, imperdiet a, venenatis vitae, justo. Nullam dictum felis eu pede mollis pretium. Integer tincidunt. Cras dapibus. Vivamus elementum semper nisi. Aenean vulputate eleifend tellus. Aenean leo ligula, porttitor eu, consequat vitae, eleifend ac, enim. Aliquam lorem ante, dapibus in, viverra quis, feugiat a, tellus. Phasellus viverra nulla ut metus varius laoreet. Quisque rutrum. Aenean imperdiet. Etiam ultricies nisi vel augue. Curabitur ullamcorper ultricies nisi. Nam eget dui. Etiam rhoncus. Maecenas tempus, tellus eget condimentum rhoncus, sem quam semper libero, sit amet adipiscing sem neque sed ipsum. Nam quam nunc, blandit vel, luctus pulvinar, hendrerit id, lorem. Maecenas nec odio et ante tincidunt tempus. Donec vitae sapien ut libero venenatis faucibus. Nullam quis ante. Etiam sit amet orci eget eros faucibus tincidunt. Duis leo. Sed fringilla mauris sit amet nibh. Donec sodales sagittis magna.",
				},
				{ text: "" },
				{ text: "" },
				{ text: "" },
				{ text: "" },
				{ text: "" },
				{ text: "" },
			],
		];

		slide.addTable(projRows2, {
			x: 0.4,
			y: 5.25,
			colW: [0.5, 1.8, 5, 0.9, 1.0, 0.95, 0.8, 1.5],
			border: { pt: 0.1, color: "818181" },
			align: "left",
			valign: "middle",
			fontFace: "Segoe UI",
			fontSize: 8,
			autoPage: true,
			autoPageRepeatHeader: true,
			autoPageLineWeight: -0.4,
			verbose: true,
		});
	}

	// SLIDE 3:
	{
		slide = pptx.addSlide({ sectionTitle: "Tables: QA" });

		let projRows = [
			[
				{ text: "id", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "First item Desc", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Impact", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Owner", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Created Date", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Due Date", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Status", options: { bold: true, fill: "1F3864", color: "ffffff" } },
				{ text: "Update", options: { bold: true, fill: "1F3864", color: "ffffff" } },
			],
			[
				{ text: "1" },
				{
					text: "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, arcu. In enim justo, rhoncus ut, imperdiet a, venenatis vitae, justo. Nullam dictum felis eu pede mollis pretium. Integer tincidunt. Cras dapibus. Vivamus elementum semper nisi. Aenean vulputate eleifend tellus. Aenean leo ligula, porttitor eu, consequat vitae, eleifend ac, enim. Aliquam lorem ante, dapibus in, viverra quis, feugiat a, tellus. Phasellus viverra nulla ut metus varius laoreet. Quisque rutrum. Aenean imperdiet. Etiam ultricies nisi vel augue. Curabitur ullamcorper ultricies nisi. Nam eget dui. Etiam rhoncus. Maecenas tempus, tellus eget condimentum rhoncus, sem quam semper libero, sit amet adipiscing sem neque sed ipsum. Nam quam nunc, blandit vel, luctus pulvinar, hendrerit id, lorem. Maecenas nec odio et ante tincidunt tempus. Donec vitae sapien ut libero venenatis faucibus. Nullam quis ante. Etiam sit amet orci eget eros faucibus tincidunt. Duis leo. Sed fringilla mauris sit amet nibh. Donec sodales sagittis magna.",
				},
				{
					text: "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Aenean commodo ligula eget dolor. Aenean massa. Cum sociis natoque penatibus et magnis dis parturient montes, nascetur ridiculus mus. Donec quam felis, ultricies nec, pellentesque eu, pretium quis, sem. Nulla consequat massa quis enim. Donec pede justo, fringilla vel, aliquet nec, vulputate eget, arcu. In enim justo, rhoncus ut, imperdiet a, venenatis vitae, justo. Nullam dictum felis eu pede mollis pretium. Integer tincidunt. Cras dapibus. Vivamus elementum semper nisi. Aenean vulputate eleifend tellus. Aenean leo ligula, porttitor eu, consequat vitae, eleifend ac, enim. Aliquam lorem ante, dapibus in, viverra quis, feugiat a, tellus. Phasellus viverra nulla ut metus varius laoreet. Quisque rutrum. Aenean imperdiet. Etiam ultricies nisi vel augue. Curabitur ullamcorper ultricies nisi. Nam eget dui. Etiam rhoncus. Maecenas tempus, tellus eget condimentum rhoncus, sem quam semper libero, sit amet adipiscing sem neque sed ipsum. Nam quam nunc, blandit vel, luctus pulvinar, hendrerit id, lorem. Maecenas nec odio et ante tincidunt tempus. Donec vitae sapien ut libero venenatis faucibus. Nullam quis ante. Etiam sit amet orci eget eros faucibus tincidunt. Duis leo. Sed fringilla mauris sit amet nibh. Donec sodales sagittis magna.",
				},
				{ text: "Aenean commodo ligula eget dolor. Aenean massa." },
				{ text: "20-10-2021" },
				{ text: "01-11-2021" },
				{ text: "Pending" },
				{ text: "24-10-2021" },
			],
		];

		slide.addTable(projRows, {
			x: 0.4,
			y: 5.25,
			colW: [0.5, 3.4, 3.4, 0.9, 1.0, 0.95, 0.8, 1.5],
			border: { pt: 0.1, color: "818181" },
			align: "left",
			valign: "middle",
			fontFace: "Segoe UI",
			fontSize: 8,
			autoPage: true,
			autoPageRepeatHeader: true,
			autoPageLineWeight: -0.4,
			verbose: true,
		});
	}
}
