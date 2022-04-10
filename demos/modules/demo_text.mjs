/**
 * NAME: demo_text.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.6.0
 * BLD.: 20210426
 */

import { BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R, LOREM_IPSUM_ENG } from "./enums.mjs";

export function genSlides_Text(pptx) {
	pptx.addSection({ title: "Text" });

	genSlide01(pptx);
	genSlide02(pptx);
	genSlide03(pptx);
	genSlide04(pptx);
	genSlide05(pptx);
	genSlide06(pptx);
}

/**
 * SLIDE 1: Text alignment, percent x/y, etc.
 * @param {PptxGenJS} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Text" });

	// Slide title
	slide.addTable([[{ text: "Text Examples: Text alignment, percent x/y, etc.", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	// Slide colors: bkgd/fore
	slide.bkgd = "030303";
	slide.color = "9F9F9F";
	// Slide notes
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html");

	// Actual Textbox shape (can have any Height, can wrap text, etc.)
	slide.addText(
		[
			{ text: "Textbox align (center/middle)", options: { fontSize: 32, breakLine: true } },
			{ text: "Character Spacing 16", options: { fontSize: 16, charSpacing: 16, breakLine: true } },
			{ text: "Transparency 50%", options: { fontSize: 16, transparency: 50 } },
		],
		{ x: 0.5, y: 0.75, w: 8.5, h: 2.5, color: "FFFFFF", fill: { color: "0000FF" }, valign: "middle", align: "center", isTextBox: true }
	);
	slide.addText(
		[
			{ text: "(left/top)", options: { fontSize: 12, breakLine: true } },
			{ text: "Textbox", options: { bold: true } },
		],
		{ x: 10, y: 0.75, w: 3.0, h: 1.0, color: "FFFFFF", fill: { color: "00B050" }, valign: "top", align: "left", margin: 15 }
	);
	slide.addText(
		[
			{ text: "Textbox", options: { breakLine: true } },
			{ text: "(right/bottom)", options: { fontSize: 12 } },
		],
		{ x: 10, y: 2.25, w: 3.0, h: 1.0, color: "FFFFFF", fill: { color: "C00000" }, valign: "bottom", align: "right", margin: 0 }
	);

	slide.addText("^ (50%/50%)", { x: "50%", y: "50%", w: 2 });

	slide.addText("Plain x/y coords", { x: 10, y: 4.35, w: 3 });

	slide.addText("Escaped chars: ' \" & < >", { x: 10, y: 5.35, w: 3 });

	slide.addText(
		[
			{ text: "Sub" },
			{ text: "Subscript", options: { subscript: true } },
			{ text: " // Super" },
			{ text: "Superscript", options: { superscript: true } },
		],
		{ x: 10, y: 6.3, w: 3.3 }
	);

	// TEST: using {option}: Add text box with multiline options:
	slide.addText(
		[
			{
				text: "word-level\nformatting",
				options: { fontSize: 32, fontFace: "Courier New", color: "99ABCC", align: "right", breakLine: true },
			},
			{ text: "...in the same textbox", options: { fontSize: 48, fontFace: "Arial", color: "FFFF00", align: "center" } },
		],
		{ x: 0.5, y: 4.3, w: 8.5, h: 2.5, margin: 0.1, fill: { color: "232323" } }
	);

	let objOptions = {
		x: 0,
		y: 7,
		w: "100%",
		h: 0.5,
		align: "center",
		fontFace: "Arial",
		fontSize: 24,
		color: "00EC23",
		bold: true,
		italic: true,
		underline: true,
		margin: 0,
		isTextBox: true,
	};
	slide.addText("Text: Arial, 24, green, bold, italic, underline, margin:0", objOptions);
}

/**
 * SLIDE 2: Multi-Line Formatting, Line Breaks, Line Spacing
 * @param {PptxGenJS} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Text" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html");
	slide.addTable(
		[[{ text: "Text Examples: Multi-Line Formatting, Line Breaks, Line Spacing", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS
	);

	// LEFT COLUMN ------------------------------------------------------------

	// 1: Multi-Line Formatting
	slide.addText("Word-Level Formatting:", { x: 0.5, y: 0.6, w: "40%", h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText(
		[
			{ text: "Courier New ", options: { fontSize: 36, fontFace: "Courier New", color: pptx.colors.ACCENT6 } },
			{ text: "36", options: { fontSize: 36, fontFace: "Courier New", color: pptx.colors.ACCENT1, breakLine: true } },
			{ text: "Arial ", options: { fontSize: 48, fontFace: "Arial", color: pptx.colors.ACCENT2 } },
			{ text: "48", options: { fontSize: 48, fontFace: "Courier New", color: pptx.colors.ACCENT1, breakLine: true } },
			{ text: "Verdana 48", options: { fontSize: 48, fontFace: "Verdana", color: pptx.colors.ACCENT3, align: "left", breakLine: true } },
			{
				text: "\nStrikethrough",
				options: { fontSize: 36, fontFace: "Arial", color: pptx.colors.ACCENT6, align: "center", strike: true, breakLine: true },
			},
			{
				text: "Underline",
				options: { fontSize: 36, fontFace: "Arial", color: pptx.colors.ACCENT2, align: "center", underline: true, breakLine: true },
			},
			{
				text: "FUN",
				options: {
					fontFace: "Arial",
					fontSize: 48,
					color: pptx.colors.ACCENT6,
					align: "center",
					underline: { style: "wavy", color: pptx.colors.ACCENT4 },
				},
			},
			{
				text: "derline",
				options: {
					fontFace: "Arial",
					fontSize: 48,
					color: pptx.colors.ACCENT4,
					align: "center",
					underline: { style: "wavy", color: pptx.colors.ACCENT5 },
					breakLine: true,
				},
			},
			{ text: " ", options: { breakLine: true } },
			{ text: "Also: ", options: { fontSize: 36, fontFace: "Arial", color: pptx.colors.ACCENT5, align: "right" } },
			{ text: "highlighted", options: { fontSize: 36, fontFace: "Arial", color: pptx.colors.ACCENT5, align: "right", highlight: "FFFF00" } },
			{ text: " text!", options: { fontSize: 36, fontFace: "Arial", color: pptx.colors.ACCENT5, align: "right" } },
		],
		{ x: 0.5, y: 1.0, w: 5.75, h: 6.0, margin: 5, fill: { color: pptx.colors.TEXT1 } }
	);

	// RIGHT COLUMN ------------------------------------------------------------

	// 1: Line-Breaks
	slide.addText("Line-Breaks:", { x: 7.0, y: 0.6, w: "40%", h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText("***Line Breaks / Multi Lines***\nFirst line\nSecond line\nThird line", {
		x: 7.0,
		y: 1.0,
		w: 5.75,
		h: 1.6,
		valign: "middle",
		align: "center",
		color: "6c6c6c",
		fontSize: 16,
		fill: "F2F2F2",
		line: { color: "C7C7C7", width: "2" },
	});

	// 2: Line-Spacing (exact)
	slide.addText("Line-Spacing (text):", { x: 7.0, y: 2.7, w: "40%", h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText(
		"lineSpacing (Exactly)\n40pt",
		{ x: 7.0, y: 3.1, w: 5.75, h: 1.17, align: "center", fill: { color: "F1F1F1" }, color: "363636", lineSpacing: 39.9 } // TEST-CASE: `lineSpacing` decimal value
	);

	// 3: Line-Spacing (multiple)
	slide.addText("lineSpacing (Multiple)\n1.5", {
		x: 7.0,
		y: 4.5,
		w: 5.75,
		h: 1.0,
		align: "center",
		fill: { color: "E6E9EC" },
		color: "363636",
		lineSpacingMultiple: 1.5,
	});

	// 4: Line-Spacing (bullets)
	slide.addText("Line-Spacing (bullets):", { x: 7.0, y: 5.6, w: "40%", h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText([{ text: "lineSpacing\n35pt", options: { fontSize: 24, bullet: true, color: "99ABCC", lineSpacing: 35 } }], {
		x: 7.0,
		y: 6.0,
		w: 5.75,
		h: 1,
		margin: [0, 0, 0, 10],
		fill: { color: "F1F1F1" },
	});
}

/**
 * SLIDE 3: Bullets
 * @param {PptxGenJS} pptx
 */
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Text" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html");
	slide.addTable([[{ text: "Text Examples: Bullets", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// LEFT COLUMN ------------------------------------------------------------

	// 1: Bullets with indent levels
	slide.addText("Bullet Indent-Levels:", { x: 0.5, y: 0.6, w: "40%", h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText(
		[
			{ text: "Root-Level    ", options: { fontSize: 32, bullet: true, color: pptx.colors.ACCENT3, indentLevel: 0 } },
			{ text: "Indent-Level 1", options: { fontSize: 32, bullet: true, color: pptx.colors.ACCENT4, indentLevel: 1 } },
			{ text: "Indent-Level 2", options: { fontSize: 32, bullet: true, color: pptx.colors.ACCENT5, indentLevel: 2 } },
			{ text: "Indent-Level 3", options: { fontSize: 32, bullet: true, color: pptx.colors.ACCENT6, indentLevel: 3 } },
		],
		{ x: 0.5, y: 1.0, w: 5.75, h: 2.4, fill: { color: "232323" } }
	);

	slide.addText("Bullet Spacing (Indentation):", { x: 0.5, y: 3.5, w: "40%", h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText(
		[
			{ text: "bullet indent: 10", options: { bullet: { indent: 10 } } },
			{ text: "bullet indent: 30", options: { bullet: { indent: 30 } } },
		],
		{ x: 0.5, y: 3.9, w: 5.75, h: 0.5, color: "393939", fontFace: "Arial", fontSize: 12, fill: { color: pptx.colors.BACKGROUND2 } }
	);

	slide.addText("Bullet Styles:", { x: 0.5, y: 4.6, w: "40%", h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText(
		[
			{ text: "style: arabicPeriod", options: { color: pptx.colors.ACCENT2, bullet: { type: "number", style: "arabicPeriod" } } },
			{ text: "style: arabicPeriod", options: { color: pptx.colors.ACCENT2, bullet: { type: "number", style: "arabicPeriod" } } },
			{
				text: "style: alphaLcPeriod",
				options: { color: pptx.colors.ACCENT5, bullet: { type: "number", style: "alphaLcPeriod" }, indentLevel: 1 },
			},
			{
				text: "style: alphaLcPeriod",
				options: { color: pptx.colors.ACCENT5, bullet: { type: "number", style: "alphaLcPeriod" }, indentLevel: 1 },
			},
			{
				text: "style: romanLcPeriod",
				options: { color: pptx.colors.ACCENT6, bullet: { type: "number", style: "romanLcPeriod" }, indentLevel: 2 },
			},
			{
				text: "style: romanLcPeriod",
				options: { color: pptx.colors.ACCENT6, bullet: { type: "number", style: "romanLcPeriod" }, indentLevel: 2 },
			},
		],
		{ x: 0.5, y: 5.0, w: 5.75, h: 2.0, fill: { color: pptx.colors.BACKGROUND2 }, color: pptx.colors.ACCENT1 }
	);

	// RIGHT COLUMN ------------------------------------------------------------

	// 1: Regular bullets
	slide.addText('Bullet "Start At" number option:', { x: 7.0, y: 0.6, w: 5.75, h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText("type:'number'\nnumberStartAt:'5'", {
		x: 7.0,
		y: 1.0,
		w: 5.75,
		h: 0.75,
		fill: { color: pptx.colors.BACKGROUND2 },
		color: pptx.colors.ACCENT6,
		fontFace: "Courier New",
		bullet: { type: "number", numberStartAt: "5" },
	});

	// 2: Bullets: Text With Line-Breaks
	slide.addText("Bullets made with line breaks:", { x: 7.0, y: 1.95, w: 5.75, h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText("Line 1\nLine 2\nLine 3", {
		x: 7.0,
		y: 2.35,
		w: 5.75,
		h: 1.0,
		color: "393939",
		fontSize: 16,
		fill: pptx.colors.BACKGROUND2,
		bullet: { type: "number" },
	});

	// 3: Bullets: Soft-Line-Breaks
	slide.addText("Bullets and soft-line-break (shift+enter):", { x: 7.0, y: 3.5, w: 5.75, h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText(
		[
			{ text: "First line", options: { bullet: true, breakLine: true } },
			{ text: "Second line", options: { bullet: true } },
			{ text: "Third line via `softBreakBefore:true`", options: { softBreakBefore: true } },
		],
		{ x: 7.0, y: 3.9, w: 5.75, h: 1.0, color: "393939", fontSize: 16, fill: pptx.colors.BACKGROUND2 }
	);

	// 3: Bullets: With custom unicode bullet characters
	slide.addText("Bullets with text objects:", { x: 7.0, y: 5.05, w: 5.75, h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText(
		[
			{ text: "`bullet: { code: '25BA' }`", options: { fontSize: 18, color: pptx.colors.ACCENT1, bullet: { code: "25BA" } } },
			{ text: "`bullet: { code: '25D1' }`", options: { fontSize: 18, color: pptx.colors.ACCENT5, bullet: { code: "25D1" } } },
			{ text: "`bullet: { code: '25CC' }`", options: { fontSize: 18, color: pptx.colors.ACCENT6, bullet: { code: "25CC" } } },
			{ text: "Mix and... ", options: { fontSize: 24, color: "FF0000", bullet: { code: "25BA" } } },
			{ text: "match formatting as well.", options: { fontSize: 16, color: "00CD00" } },
		],
		{ x: 7.0, y: 5.5, w: 5.75, h: 1.5, fontFace: "Arial", fill: pptx.colors.BACKGROUND2 }
	);
}

/**
 * SLIDE 4: Hyperlinks, Text Shadow, Text Outline, Text Glow
 * @param {PptxGenJS} pptx
 */
function genSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Text" });

	slide.addTable(
		[[{ text: "Text Examples: Hyperlinks, Tab Stops, Text Effects: Shadow, Outline, and Glow", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]],
		BASE_TABLE_OPTS
	);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html");

	// 1: TOP_ROW: Hyperlinks
	slide.addText("Hyperlinks:", { x: 0.5, y: 0.6, w: "90%", h: 0.3, margin: 0.123, color: pptx.colors.ACCENT1 });
	slide.addText(
		[{ text: "Link with Tooltip", options: { hyperlink: { url: "https://github.com/gitbrent/pptxgenjs", tooltip: "Visit Homepage" } } }],
		{ x: 0.5, y: 1.0, w: 2.5, h: 0.6, margin: 10, fill: { color: "F1F1F1" }, fontSize: 14, align: "center" }
	);
	slide.addText([{ text: "Link without Tooltip", options: { hyperlink: { url: "https://github.com/gitbrent" } } }], {
		x: 3.78,
		y: 1.0,
		w: 2.5,
		h: 0.6,
		margin: 10,
		fill: { color: "F1F1F1" },
		fontSize: 14,
		align: "center",
	});
	slide.addText([{ text: "Link with custom color", options: { hyperlink: { url: "https://github.com/gitbrent" }, color: "EE40EE" } }], {
		x: 7.05,
		y: 1.0,
		w: 2.5,
		h: 0.6,
		margin: 10,
		fill: { color: "F1F1F1" },
		fontSize: 14,
		align: "center",
	});
	slide.addText([{ text: "Link to Slide #5", options: { hyperlink: { slide: 5 } } }], {
		x: 10.33,
		y: 1.0,
		w: 2.5,
		h: 0.6,
		margin: 10,
		fill: { color: "E2F0D9" },
		fontSize: 14,
		align: "center",
	});

	// 2: CTR_ROW: Tab Stops: Set tab points (inches), then use "\t" to add tab characters in your text string
	slide.addText("Tab Stops:", { x: 0.5, y: 2.1, w: 12.0, h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText([{ text: "text...\tTab1\tTab2\tTab3", options: { tabStops: [{ position: 1 }, { position: 3 }, { position: 7 }] } }], {
		x: 0.5,
		y: 2.5,
		w: 12.3,
		h: 0.6,
		fill: { color: pptx.colors.BACKGROUND2 },
	});
	slide.addText(
		"// Code for example above\n" +
			"{\n" +
			"  text: 'text...\\tTab1\\tTab2\\tTab3',\n" +
			"  options: { tabStops: [{ position: 1 }, { position: 3 }, { position: 7 }] },\n" +
			"};",
		{ x: 0.5, y: 3.3, w: 12.3, h: 2.0, fontFace: "Courier", fontSize: 13, fill: { color: "D1E1F1" }, color: "363636" }
	);

	// 3a: BTM_ROW: Text Effects: Outline
	slide.addText("Text Outline:", { x: 0.5, y: 5.8, w: 3.0, h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText("size:2", {
		x: 0.5,
		y: 6.2,
		w: 3.0,
		h: 1.1,
		fontSize: 72,
		bold: true,
		color: pptx.colors.ACCENT1,
		outline: { size: 2, color: pptx.colors.ACCENT4 },
	});

	// 3b: Text Effects: Glow
	slide.addText("Text Glow:", { x: 3.9, y: 5.8, w: 5.0, h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText("size:10", {
		x: 3.9,
		y: 6.2,
		w: 3.0,
		h: 1.1,
		fontSize: 72,
		color: pptx.colors.ACCENT1,
		glow: { size: 10, opacity: 0.25, color: pptx.colors.ACCENT2 },
	});

	// 3c: Text Effects: Shadow
	let shadowOpts = { type: "outer", color: "696969", blur: 3, offset: 10, angle: 45, opacity: 0.6 };
	slide.addText("Text Shadow:", { x: 7.5, y: 5.8, w: 5.0, h: 0.3, margin: 0, color: pptx.colors.ACCENT1 });
	slide.addText("type:outer, offset:10, blur:3", { x: 7.5, y: 6.2, w: 5.5, h: 1.1, fontSize: 32, color: "0088cc", shadow: shadowOpts });
}

/**
 * SLIDE 5: Text Fit: Shrink/Resize
 * @param {PptxGenJS} pptx
 */
function genSlide05(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Text" });

	slide.addTable([[{ text: "Text Examples: Text Fit", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html");

	slide.addText(LOREM_IPSUM_ENG.substring(0, 1200), { x: 0.5, y: 1.3, w: 4, h: 4, fontSize: 12, fit: "none" });
	slide.addText(LOREM_IPSUM_ENG.substring(0, 1200), { x: 4.5, y: 1.3, w: 4, h: 4, fontSize: 12, fit: "shrink" });
	slide.addText(LOREM_IPSUM_ENG.substring(0, 1200), { x: 8.5, y: 1.3, w: 4, h: 4, fontSize: 12, fit: "resize" });

	// titles last so they overlay the overflowing text from above
	slide.addText("fit:'none'  ", { x: 0.5, y: 0.6, w: 4, h: 0.3, color: pptx.colors.ACCENT1, fill: { color: "ffffff" } });
	slide.addText("fit:'shrink'", { x: 4.5, y: 0.6, w: 4, h: 0.3, color: pptx.colors.ACCENT1, fill: { color: "ffffff" } });
	slide.addText("fit:'resize'", { x: 8.5, y: 0.6, w: 4, h: 0.3, color: pptx.colors.ACCENT1, fill: { color: "ffffff" } });

	slide.addText(
		[
			{ text: "NOTE", options: { fontSize: 16, bold: true, breakLine: true } },
			{
				text: "- both 'Shrink' and 'Resize' are only applied once text is editted or the shape is resized after creation using PowerPoint/Keynote/et al.",
				options: { breakLine: true },
			},
			{
				text: "- PowerPoint calculates a scaling factor and applies it dynamically when a shape is updated - something that cannot be triggered by PptxGenJS",
				options: { breakLine: true },
			},
			{
				text: "- the textboxes above have their shrink & resize props set already, just add a space or resize them to trigger shrink and resize behavior",
			},
		],
		{ x: 0.5, y: 6.0, w: 12, h: 1.1, margin: 10, fontSize: 12, color: "393939", fill: { color: "fffccc" } }
	);
}

/**
 * SLIDE 6: Scheme Colors
 * @param {PptxGenJS} pptx
 */
function genSlide06(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Text" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html");
	slide.addTable([[{ text: "Text Examples: Scheme Colors", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	// MISC ------------------------------------------------------------

	slide.addText("TEXT1 on BACKGROUND2", { x: 0.5, y: 0.7, w: 6.0, h: 2.0, color: pptx.colors.TEXT1, fill: { color: pptx.colors.BACKGROUND2 } });
	slide.addText("TEXT2 on BACKGROUND1", { x: 7.0, y: 0.7, w: 6.0, h: 2.0, color: pptx.colors.TEXT2, fill: { color: pptx.colors.BACKGROUND1 } });

	slide.addText("ACCENT1", { x: 0.5, y: 3.25, w: 6.0, h: 1.0, color: "FFFFFF", fill: { color: pptx.colors.ACCENT1 } });
	slide.addText("ACCENT2", { x: 7.0, y: 3.25, w: 6.0, h: 1.0, color: "FFFFFF", fill: { color: pptx.colors.ACCENT2 } });
	slide.addText("ACCENT3", { x: 0.5, y: 4.7, w: 6.0, h: 1.0, color: "FFFFFF", fill: { color: pptx.colors.ACCENT3 } });
	slide.addText("ACCENT4", { x: 7.0, y: 4.7, w: 6.0, h: 1.0, color: "FFFFFF", fill: { color: pptx.colors.ACCENT4 } });
	slide.addText("ACCENT5", { x: 0.5, y: 6.15, w: 6.0, h: 1.0, color: "FFFFFF", fill: { color: pptx.colors.ACCENT5 } });
	slide.addText("ACCENT6", { x: 7.0, y: 6.15, w: 6.0, h: 1.0, color: "FFFFFF", fill: { color: pptx.colors.ACCENT6 } });

	// NEGATIVE TEST:
	//slide.addText("NEGTEST / NEGTEST", { x:0.5, y:0.5, w:'40%', h:0.38, color:pptx.colors.NEGTEST01, fill:{color:pptx.colors.NEGTEST02} });
}
