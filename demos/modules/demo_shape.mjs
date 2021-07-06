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
	genSlide03(pptx);
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
	slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
		x: 10,
		y: 2.5,
		w: 3.0,
		h: 1.5,
		rectRadius: 1,
		fill: { color: "00FF00" },
		line: "000000",
		lineSize: 1,
	}); // TEST: DEPRECATED: `fill`,`line`,`lineSize`
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
	slide.addText("ROUNDED-RECTANGLE\ndashType:dash\nrectRadius:1", {
		shape: pptx.shapes.ROUNDED_RECTANGLE,
		x: 10,
		y: 2.5,
		w: 3.0,
		h: 1.5,
		fill: { color: "00FF00" },
		align: "center",
		fontSize: 14,
		line: { color: "000000", size: 1, dashType: "dash" },
		rectRadius: 1,
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

function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Shapes" });

	slide.addTable([[{ text: "Shape Examples 3: Custom Geometry: Text & Shapes", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-shapes.html");

	// TODO: add messagebox with:
	/*
		I've implemented this by using a similar spec to the one that uses `svg-points`.
		The path or contour of the custom geometry is declared under the property points of the ShapeProps object.
		With this implementation we are supporting all the custom geometry rules: moveTo, lnTo, arcTo, cubicBezTo, quadBezTo and close.

		A translation of an svg path to a custom geometry could be achieved by using the svg-points package and adding a custom translation between the arcs.
		The svg arc is described by the variables x, y, rx, ry, xAxisRotation, largeArcFlag and sweepFlag.
		On the other side the pptx freeform arc is described by x, y, hR, wR, stAng, swAng.
		In order to add some sort of translation between svg-path and a custom geometry points array we should create a translation between those two representations of the arc.
	*/

	// BROKEN HEART SVG
	/*
		<defs>
			<g id="left">
				<path d="M0 0
					c-25,-80 -50,-80 -100,-130
					a70,70 -45 0,1 100,-100
					l25 60 l-55 50 l40 55 l-10 65
					" />
			</g>
			<g id="right">
				<path d="M0 -230
					a70,70 45 0,1 100,100
					c-50,50 -75,50 -100,130
					l10 -65 l-40 -55 l55 -50 l-25 -60
					" />
			</g>
		</defs>
	*/
	/*
	slide.addShape(pptx.shapes.CUSTOM_GEOMETRY, {
		x: 1,
		y: 0.8,
		w: 285 / 100,
		h: 285 / 100,
		fill: { color: "FF0000" },
		line: { color: "CC0000", width: 1 },
		points: [
			{ x: 0, y: 0 },
			{ x: 100 / 100, y: -100 / 100, curve: { type: "arc", hR: 0.7, wR: 0.7, stAng: -45, swAng: -45 } },
			{ x: 25 / 100, y: 60 / 100 },
			{ x: 55 / 100, y: 50 / 100 },
			{ x: 40 / 100, y: 55 / 100 },
			{ x: -10 / 100, y: 65 / 100 },
			{ close: true },
		],
	});
	*/

	// EXAMPLE 1:
	slide.addShape(pptx.shapes.CUSTOM_GEOMETRY, {
		x: 1,
		y: 0.8,
		w: 3.0,
		h: 1.5,
		fill: { color: "00FF00" },
		line: { color: "000000", width: 1 },
		points: [
			{ x: 0, y: 0.75 },
			{ x: 0.5, y: 0 },
			{ x: 1, y: 0.3 },
			{ x: 1.5, y: 0 },
			{ x: 2, y: 0.3 },
			{ x: 2.5, y: 0.2 },
			{ x: 3, y: 0.1 },
			{ curve: { type: "arc", hR: 0.5, wR: 0.5, stAng: 0, swAng: 90 } },
			{ x: 0.5, y: 1.5, curve: { type: "quadratic", x1: 2.5, y1: 1.7 } },
			{ close: true },
		],
	});

	// EXAMPLE 2:
	slide.addText("CUSTOM-GEOMETRY", {
		shape: pptx.shapes.CUSTOM_GEOMETRY,
		x: 10,
		y: 0.8,
		w: 3.0,
		h: 1.5,
		fill: { color: "00FF00" },
		line: { color: "000000", width: 1 },
		points: [
			{ x: 0, y: 0.75 },
			{ x: 0.5, y: 0 },
			{ x: 1, y: 0.3 },
			{ x: 1.5, y: 0 },
			{ x: 2, y: 0.3 },
			{ x: 2.5, y: 0.2 },
			{ x: 3, y: 0.1 },
			{ curve: { type: "arc", hR: 0.5, wR: 0.5, stAng: 0, swAng: 90 } },
			{ x: 0.5, y: 1.5, curve: { type: "quadratic", x1: 2.5, y1: 1.7 } },
			{ close: true },
		],
	});
}
