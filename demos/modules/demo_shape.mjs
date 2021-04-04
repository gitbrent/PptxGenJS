/**
 * NAME: demo_shapes.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.5.0
 * BLD.: 20210401
 */

import { BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R } from "./enums.mjs";

export function genSlides_Shape(pptx) {
	pptx.addSection({ title: "Shapes" });

	genSlide01(pptx);
	genSlide02(pptx);
}

/**
 * SLIDE 1: Misc Shape Types (no text)
 * @param {PptxGenJS} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Shapes" });

	slide.addTable([[{ text: "Shape Examples 1: Misc Shape Types (no text)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-shapes.html");

	slide.addShape(pptx.shapes.RECTANGLE, { x: 0.5, y: 0.8, w: 1.5, h: 3.0, fill: { color: "FF0000" }, line: { type: "none" } });
	slide.addShape(pptx.shapes.RECTANGLE, { x: 3.0, y: 0.7, w: 1.5, h: 3.0, fill: { color: "F38E00" }, rotate: 45 });
	slide.addShape(pptx.shapes.OVAL, { x: 5.4, y: 0.8, w: 3.0, h: 1.5, fill: { type: "solid", color: "0088CC" } });
	slide.addShape(pptx.shapes.OVAL, { x: 7.7, y: 1.4, w: 3.0, h: 1.5, fill: { color: "FF00CC" }, rotate: 90 }); // TEST: no type
	slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 10, y: 2.5, w: 3.0, h: 1.5, r: 0.2, fill: { color: "00FF00" }, line: "000000", lineSize: 1 }); // TEST: DEPRECATED: `fill`,`line`,`lineSize`
	slide.addShape(pptx.shapes.ARC, { x: 10.75, y: 0.8, w: 1.5, h: 1.5, fill: { color: "0088CC" }, angleRange: [45, 315] });
	//
	slide.addShape(pptx.shapes.LINE, { x: 4.2, y: 4.4, w: 5.0, h: 0.0, line: { color: "FF0000", width: 1, dashType: "lgDash" } });
	slide.addShape(pptx.shapes.LINE, { x: 4.2, y: 4.8, w: 5.0, h: 0.0, line: { color: "FF0000", width: 2, dashType: "dashDot" }, lineHead: "arrow" }); // TEST: DEPRECATED: lineHead
	slide.addShape(pptx.shapes.LINE, { x: 4.2, y: 5.2, w: 5.0, h: 0.0, line: { color: "FF0000", width: 3, endArrowType: "triangle" } });
	slide.addShape(pptx.shapes.LINE, {
		x: 4.2,
		y: 5.6,
		w: 5.0,
		h: 0.0,
		line: { color: "FF0000", width: 4, beginArrowType: "diamond", endArrowType: "oval" },
	});
	slide.addShape(pptx.shapes.LINE, { x: 5.7, y: 3.3, w: 2.5, h: 0.0, line: { width: 1 }, rotate: 360 - 45 }); // DIAGONAL Line // TEST:no line.color
	//
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE, {
		x: 0.4,
		y: 4.3,
		w: 6.0,
		h: 3.0,
		fill: { color: "0088CC" },
		line: { color: "000000", width: 3 },
		shapeName: "First Right Triangle",
	});
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE, {
		x: 7.0,
		y: 4.3,
		w: 6.0,
		h: 3.0,
		fill: { color: "0088CC" },
		line: { color: "000000", width: 2 },
		flipH: true,
	});
}

/**
 * SLIDE 2: Misc Shape Types with Text
 * @param {PptxGenJS} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Shapes" });

	slide.addTable([[{ text: "Shape Examples 2: Misc Shape Types (with text)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-shapes.html");

	slide.addText("RECTANGLE", {
		shape: pptx.shapes.RECTANGLE,
		x: 0.5,
		y: 0.8,
		w: 1.5,
		h: 3.0,
		fill: { color: "FF0000" },
		align: "center",
		fontSize: 14,
	});
	slide.addText("RECTANGLE (rotate:45)", {
		shape: pptx.shapes.RECTANGLE,
		x: 3.0,
		y: 0.7,
		w: 1.5,
		h: 3.0,
		fill: { color: "F38E00" },
		rotate: 45,
		align: "center",
		fontSize: 14,
	});
	slide.addText("OVAL (transparency:50)", {
		shape: pptx.shapes.OVAL,
		x: 5.4,
		y: 0.8,
		w: 3.0,
		h: 1.5,
		fill: { type: "solid", color: "0088CC", transparency: 50 },
		align: "center",
		fontSize: 14,
	});
	// TEST: DEPRECATED: `alpha`
	slide.addText("OVAL (rotate:90, transparency:75)", {
		shape: pptx.shapes.OVAL,
		x: 7.7,
		y: 1.4,
		w: 3.0,
		h: 1.5,
		fill: { type: "solid", color: "FF00CC", alpha: 75 },
		rotate: 90,
		align: "center",
		fontSize: 14,
	});
	slide.addText("ROUNDED-RECTANGLE\ndashType:dash\nrectRadius:10", {
		shape: pptx.shapes.ROUNDED_RECTANGLE,
		x: 10,
		y: 2.5,
		w: 3.0,
		h: 1.5,
		r: 0.2,
		fill: { color: "00FF00" },
		align: "center",
		fontSize: 14,
		line: { color: "000000", size: 1, dashType: "dash" },
		rectRadius: 10,
	});
	slide.addText("ARC", {
		shape: pptx.shapes.ARC,
		x: 10.75,
		y: 0.8,
		w: 1.5,
		h: 1.5,
		fill: { color: "0088CC" },
		angleRange: [45, 315],
		line: { color: "002244", width: 1 },
		fontSize: 14,
	});
	//
	slide.addText("LINE size=1", {
		shape: pptx.shapes.LINE,
		align: "center",
		x: 4.15,
		y: 4.4,
		w: 5,
		h: 0,
		line: { color: "FF0000", width: 1, dashType: "lgDash" },
	});
	slide.addText("LINE size=2", {
		shape: pptx.shapes.LINE,
		align: "left",
		x: 4.15,
		y: 4.8,
		w: 5,
		h: 0,
		line: { color: "FF0000", width: 2, dashType: "dashDot", endArrowType: "arrow" },
	});
	slide.addText("LINE size=3", {
		shape: pptx.shapes.LINE,
		align: "right",
		x: 4.15,
		y: 5.2,
		w: 5,
		h: 0,
		line: { color: "FF0000", width: 3, beginArrowType: "triangle" },
	});
	slide.addText("LINE size=4", {
		shape: pptx.shapes.LINE,
		x: 4.15,
		y: 5.6,
		w: 5,
		h: 0,
		line: { color: "FF0000", width: 4, beginArrowType: "diamond", endArrowType: "oval", transparency: 50 },
	});
	slide.addText("DIAGONAL", { shape: pptx.shapes.LINE, valign: "bottom", x: 5.7, y: 3.3, w: 2.5, h: 0, line: { width: 2 }, rotate: 360 - 45 }); // TEST: (missing `line.color`)
	//
	slide.addText("RIGHT-TRIANGLE", {
		shape: pptx.shapes.RIGHT_TRIANGLE,
		align: "center",
		x: 0.4,
		y: 4.3,
		w: 6,
		h: 3,
		fill: { color: "0088CC" },
		line: { color: "000000", width: 3 },
	});
	slide.addText("HYPERLINK-SHAPE", {
		shape: pptx.shapes.RIGHT_TRIANGLE,
		align: "center",
		x: 7.0,
		y: 4.3,
		w: 6,
		h: 3,
		fill: { color: "0088CC" },
		line: { color: "000000", width: 2 },
		flipH: true,
		hyperlink: { url: "https://github.com/gitbrent/pptxgenjs", tooltip: "Visit Homepage" },
	});
}
