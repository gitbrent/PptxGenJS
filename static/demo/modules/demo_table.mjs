/**
 * NAME: demo_tables.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.6.0
 * BLD.: 20210404
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
	genSlide07(pptx);
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

	let tabOpts3 = {
		x: 0.5,
		y: 5.15,
		w: 6.25,
		h: 2,
		margin: 0.25,
		align: "center",
		valign: "middle",
		fontSize: 16,
		border: { pt: "1", color: "c7c7c7" },
		fill: { color: "F1F1F1" },
	};
	let arrTabRows3 = [
		[
			{ text: "A1\nA2\nA3", options: { rowspan: 3, fill: { color: "FFFCCC" } } },
			{ text: "B1\nB2", options: { rowspan: 2, fill: { color: "FFFCCC" } } },
			"C1",
		],
		["C2"],
		[{ text: "B3 -> C3", options: { colspan: 2, fill: { color: "99FFCC" } } }],
	];
	slide.addTable(arrTabRows3, tabOpts3);

	let tabOpts4 = {
		x: 7.4,
		y: 5.15,
		w: 5.5,
		h: 2,
		margin: 0,
		align: "center",
		valign: "middle",
		fontSize: 16,
		border: { pt: "1", color: "c7c7c7" },
		fill: { color: "f2f9fc" },
	};
	let arrTabRows4 = [
		[
			"A1",
			{ text: "B1\nB2", options: { rowspan: 2, fill: { color: "FFFCCC" } } },
			{ text: "C1\nC2\nC3", options: { rowspan: 3, fill: { color: "FFFCCC" } } },
		],
		["A2"],
		[{ text: "A3 -> B3", options: { colspan: 2, fill: { color: "99FFCC" } } }],
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
	slide.addTable([["margin:[0,0,0,20]"]], { x: 2.5, y: 1.1, margin: [0, 0, 0, 20], w: 2.0, fill: "FFFCCC", align: "right" });
	slide.addTable([["margin:5"]], { x: 5.5, y: 1.1, margin: 5, w: 1.0, fill: pptx.SchemeColor.background2 });
	slide.addTable([["margin:[40,5,5,20]"]], { x: 7.5, y: 1.1, margin: [40, 5, 5, 20], w: 2.2, fill: "F1F1F1" });
	slide.addTable([["margin:[30,5,5,30]"]], { x: 10.5, y: 1.1, margin: [30, 5, 5, 30], w: 2.2, fill: "F1F1F1" });

	slide.addTable(
		[
			[
				{ text: "no border and number zero", options: { margin: 5 } },
				{ text: 0, options: { margin: 5 } },
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
			{ text: 3, options: { color: "0088CC", align: "left" } },
		],
	];
	slide.addTable(arrTextObjects, { x: 0.5, y: 2.7, w: 12.25, margin: 7, fill: { color: "F1F1F1" }, border: { pt: 1, color: "696969" } });

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

	// Invalid char check
	optsSub.y = 6.1;
	slide.addText("Escaped Invalid Chars:", optsSub);
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

	let arrRows = [];
	let arrText = [];
	arrRows.push([
		{ text: "ID#", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
		{ text: "First Name", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
		{ text: "Lorum Ipsum", options: { fill: "0088cc", color: "ffffff", valign: "middle" } },
	]);
	TABLE_NAMES_F.forEach((name, idx) => {
		let strText = idx == 0 ? LOREM_IPSUM.substring(0, 100) : LOREM_IPSUM.substring(idx * 100, idx * 200);
		arrRows.push([idx, name, strText]);
		arrText.push([strText]);
	});

	let arrRowsHead1 = [];
	arrRows.forEach((row, idx) => {
		if (idx < 6) arrRowsHead1.push(row);
	});
	let arrRowsHead2 = [[{ text: "Title Header", options: { fill: "0088cc", color: "ffffff", align: "center", bold: true, colspan: 3, colW: 4 } }]];
	arrRows.forEach((row, idx) => {
		if (idx < 6) arrRowsHead2.push(row);
	});

	pptx.addSection({ title: "Tables: Auto-Paging" });
	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html");
	slide.addText(
		[
			{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
			{ text: "Auto-Paging Example", options: DEMO_TITLE_OPTS },
		],
		{ x: 0.5, y: 0.13, w: "90%" }
	);
	slide.addTable(arrRows, { x: 0.5, y: 0.6, colW: [0.75, 1.75, 10], margin: 2, border: { color: "CFCFCF" }, autoPage: true });

	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging" });
	slide.addText(
		[
			{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
			{ text: "Smaller Table Area", options: DEMO_TITLE_OPTS },
		],
		{ x: 0.5, y: 0.13, w: "90%" }
	);
	slide.addTable(arrRows, { x: 3.0, y: 0.6, colW: [0.75, 1.75, 7], margin: 5, border: { color: "CFCFCF" }, autoPage: true });

	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging" });
	slide.addText(
		[
			{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
			{ text: "Test: Correct starting Y location upon paging", options: DEMO_TITLE_OPTS },
		],
		{ x: 0.5, y: 0.13, w: "90%" }
	);
	slide.addTable(arrRows, { x: 3.0, y: 4.0, colW: [0.75, 1.75, 7], margin: 5, border: { color: "CFCFCF" }, fontFace: "Arial", autoPage: true });

	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging" });
	slide.addText(
		[
			{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
			{ text: "Test: `{ autoPageSlideStartY: 1.5 }`", options: DEMO_TITLE_OPTS },
		],
		{ x: 0.5, y: 0.13, w: "90%" }
	);
	slide.addTable(arrRows, {
		x: 3.0,
		y: 4.0,
		colW: [0.75, 1.75, 7],
		margin: 5,
		border: { color: "CFCFCF" },
		autoPage: true,
		autoPageSlideStartY: 1.5,
	});

	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging", masterName: "MASTER_PLAIN" });
	slide.addText(
		[
			{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
			{ text: "Master Page with Auto-Paging", options: DEMO_TITLE_OPTS },
		],
		{ x: 0.5, y: 0.13, w: "90%" }
	);
	slide.addTable(arrRows, { x: 1.0, y: 0.6, colW: [0.75, 1.75, 7], margin: 5, border: { color: "CFCFCF" }, autoPage: true });

	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging" });
	slide.addText(
		[
			{ text: "Table Examples: ", options: DEMO_TITLE_TEXT },
			{ text: "Auto-Paging Disabled", options: DEMO_TITLE_OPTS },
		],
		{ x: 0.5, y: 0.13, w: "90%" }
	);
	slide.addTable(arrRows, { x: 1.0, y: 0.6, colW: [0.75, 1.75, 7], margin: 5, border: { color: "CFCFCF" } }); // Negative-Test: no `autoPage:false`

	// `autoPageRepeatHeader` option demos
	pptx.addSection({ title: "Tables: Auto-Paging Repeat Header" });
	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging Repeat Header" });
	slide.addText(
		[
			{ text: "Table Examples: autoPageHeaderRows", options: DEMO_TITLE_TEXTBK },
			{ text: "no autoPageHeaderRows", options: DEMO_TITLE_OPTS },
		],
		{ x: 0.23, y: 0.13, w: 4, h: 0.4 }
	);
	slide.addTable(arrRowsHead1, {
		x: 0.23,
		y: 0.6,
		colW: [0.5, 1.0, 2.5],
		margin: 5,
		border: { color: "CFCFCF" },
		autoPage: true,
		autoPageRepeatHeader: true,
		autoPageSlideStartY: 0.6,
	});

	slide.addText(
		[
			{ text: "Table Examples: autoPageHeaderRows", options: DEMO_TITLE_TEXTBK },
			{ text: "autoPageHeaderRows:1", options: DEMO_TITLE_OPTS },
		],
		{ x: 4.75, y: 0.13, w: 4, h: 0.4 }
	);
	slide.addTable(arrRowsHead1, {
		x: 4.75,
		y: 0.6,
		colW: [0.5, 1.0, 2.5],
		margin: 5,
		border: { color: "CFCFCF" },
		autoPage: true,
		autoPageRepeatHeader: true,
		autoPageHeaderRows: 1,
		autoPageSlideStartY: 0.6,
	});

	slide.addText(
		[
			{ text: "Table Examples: autoPageHeaderRows", options: DEMO_TITLE_TEXTBK },
			{ text: "autoPageHeaderRows:2", options: DEMO_TITLE_OPTS },
		],
		{ x: 9.1, y: 0.13, w: 4, h: 0.4 }
	);
	slide.addTable(arrRowsHead2, {
		x: 9.1,
		y: 0.6,
		colW: [0.5, 1.0, 2.5],
		margin: 5,
		border: { color: "CFCFCF" },
		autoPage: true,
		autoPageRepeatHeader: true,
		autoPageHeaderRows: 2,
		autoPageSlideStartY: 0.6,
	});

	// autoPageLineWeight option demos
	pptx.addSection({ title: "Tables: Auto-Paging LineWeight" });
	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging LineWeight" });
	slide.addText(
		[
			{ text: "Table Examples: Line Weight Options", options: DEMO_TITLE_TEXTBK },
			{ text: "autoPageLineWeight:0.0", options: DEMO_TITLE_OPTS },
		],
		{ x: 0.23, y: 0.13, w: 4, h: 0.4 }
	);
	slide.addTable(arrText, { x: 0.23, y: 0.6, w: 4, margin: 5, border: { color: "CFCFCF" }, autoPage: true, autoPageLineWeight: 0.0 });

	slide.addText(
		[
			{ text: "Table Examples: Line Weight Options", options: DEMO_TITLE_TEXTBK },
			{ text: "autoPageLineWeight:0.5", options: DEMO_TITLE_OPTS },
		],
		{ x: 4.75, y: 0.13, w: 4, h: 0.4 }
	);
	slide.addTable(arrText, { x: 4.75, y: 0.6, w: 4, margin: 5, border: { color: "CFCFCF" }, autoPage: true, autoPageLineWeight: 0.5 });

	slide.addText(
		[
			{ text: "Table Examples: Line Weight Options", options: DEMO_TITLE_TEXTBK },
			{ text: "autoPageLineWeight:-0.5", options: DEMO_TITLE_OPTS },
		],
		{ x: 9.1, y: 0.13, w: 4, h: 0.4 }
	);
	slide.addTable(arrText, { x: 9.1, y: 0.6, w: 4, margin: 5, border: { color: "CFCFCF" }, autoPage: true, autoPageLineWeight: -0.5 });

	// autoPageCharWeight option demos
	pptx.addSection({ title: "Tables: Auto-Paging CharWeight" });
	slide = pptx.addSlide({ sectionTitle: "Tables: Auto-Paging CharWeight" });
	slide.addText(
		[
			{ text: "Table Examples: Char Weight Options", options: DEMO_TITLE_TEXTBK },
			{ text: "autoPageCharWeight:0.0", options: DEMO_TITLE_OPTS },
		],
		{ x: 0.23, y: 0.13, w: 4, h: 0.4 }
	);
	slide.addTable(arrText, { x: 0.23, y: 0.6, w: 4, margin: 5, border: { color: "CFCFCF" }, autoPage: true, autoPageCharWeight: 0.0 });

	slide.addText(
		[
			{ text: "Table Examples: Char Weight Options", options: DEMO_TITLE_TEXTBK },
			{ text: "autoPageCharWeight:0.25", options: DEMO_TITLE_OPTS },
		],
		{ x: 4.75, y: 0.13, w: 4, h: 0.4 }
	);
	slide.addTable(arrText, { x: 4.75, y: 0.6, w: 4, margin: 5, border: { color: "CFCFCF" }, autoPage: true, autoPageCharWeight: 0.25 });

	slide.addText(
		[
			{ text: "Table Examples: Char Weight Options", options: DEMO_TITLE_TEXTBK },
			{ text: "autoPageCharWeight:-0.25", options: DEMO_TITLE_OPTS },
		],
		{ x: 9.1, y: 0.13, w: 4, h: 0.4 }
	);
	slide.addTable(arrText, { x: 9.1, y: 0.6, w: 4, margin: 5, border: { color: "CFCFCF" }, autoPage: true, autoPageCharWeight: -0.25 });
}
