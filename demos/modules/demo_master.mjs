/**
 * NAME: demo_master.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.5.0
 * BLD.: 20210401
 */

import { IMAGE_PATHS } from "./enums.mjs";

export function genSlides_Master(pptx) {
	pptx.addSection({ title: "Masters" });

	genSlide01(pptx);
	genSlide02(pptx);
	genSlide03(pptx);
	genSlide04(pptx);
	genSlide05(pptx);
	genSlide06(pptx);
	//genSlide07(pptx);
}

/**
 * SLIDE 1:
 * @param {PptxGenJS} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ masterName: "TITLE_SLIDE", sectionTitle: "Masters" });
	//let slide1 = pptx.addSlide({masterName:'TITLE_SLIDE', sectionTitle:'FAILTEST'}); // TEST: Should show console warning ("title not found")
	slide.addNotes("Master name: `TITLE_SLIDE`\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
}

/**
 * SLIDE 2:
 * @param {PptxGenJS} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ masterName: "MASTER_SLIDE", sectionTitle: "Masters" });
	slide.addNotes("Master name: `MASTER_SLIDE`\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
}

/**
 * SLIDE 3:
 * @param {PptxGenJS} pptx
 */
function genSlide03(pptx) {
	let slide = pptx.addSlide({ masterName: "MASTER_SLIDE", sectionTitle: "Masters" });
	slide.addNotes("Master name: `MASTER_SLIDE` using pre-filled placeholders\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
	slide.addText("Text Placeholder", { placeholder: "header" });
	slide.addText(
		[
			{ text: "Pre-filled placeholder bullets", options: { bullet: true, valign: "top" } },
			{ text: "Add any text, charts, whatever", options: { bullet: true, indentLevel: 1, color: "0000AB" } },
			{ text: "Check out the online API docs for more", options: { bullet: true, indentLevel: 2, color: "0000AB" } },
		],
		{ placeholder: "body", valign: "top" }
	);
}

/**
 * SLIDE 4:
 * @param {PptxGenJS} pptx
 */
function genSlide04(pptx) {
	let slide = pptx.addSlide({ masterName: "MASTER_SLIDE", sectionTitle: "Masters" });
	slide.addNotes("Master name: `MASTER_SLIDE` using pre-filled placeholders\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
	slide.addText("Image Placeholder", { placeholder: "header" });
	slide.addImage({
		placeholder: "body",
		path: IMAGE_PATHS.starlabsBkgd.path,
		w: 12.0,
		h: 5.25,
	});
}

/**
 * SLIDE 5:
 * @param {PptxGenJS} pptx
 */
function genSlide05(pptx) {
	let dataChartPieLocs = [
		{
			name: "Location",
			labels: ["CN", "DE", "GB", "MX", "JP", "IN", "US"],
			values: [69, 35, 40, 85, 38, 99, 101],
		},
	];
	let slide = pptx.addSlide({ masterName: "MASTER_SLIDE", sectionTitle: "Masters" });
	slide.addNotes("Master name: `MASTER_SLIDE` using pre-filled placeholders\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
	slide.addText("Chart Placeholder", { placeholder: "header" });
	slide.addChart(pptx.charts.PIE, dataChartPieLocs, { showLegend: true, legendPos: "l", placeholder: "body" });
}

/**
 * SLIDE 6:
 * @param {PptxGenJS} pptx
 */
function genSlide06(pptx) {
	let slide = pptx.addSlide({ masterName: "THANKS_SLIDE", sectionTitle: "Masters" });
	slide.addNotes("Master name: `THANKS_SLIDE`\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html");
	slide.addText("Thank You!", { placeholder: "thanksText" });
	//slide.addText('github.com/gitbrent', { placeholder:'body' });
}

/**
 * SLIDE 7: LEGACY-TEST-ONLY: To check deprecated functionality
 * @param {PptxGenJS} pptx
 */
function genSlide07(pptx) {
	if (pptx.masters && Object.keys(pptx.masters).length > 0) {
		let slide1 = pptx.addSlide(pptx.masters.TITLE_SLIDE);
		let slide2 = pptx.addSlide(pptx.masters.MASTER_SLIDE);
		let slide3 = pptx.addSlide(pptx.masters.THANKS_SLIDE);

		let slide4 = pptx.addSlide(pptx.masters.TITLE_SLIDE, { bkgd: "0088CC", slideNumber: { x: "50%", y: "90%", color: "0088CC" } });
		let slide5 = pptx.addSlide(pptx.masters.MASTER_SLIDE, {
			bkgd: { path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/images/title_bkgd_alt.jpg" },
		});
		let slide6 = pptx.addSlide(pptx.masters.THANKS_SLIDE, { bkgd: "ffab33" });
		//let slide7 = pptx.addSlide( pptx.masters.LEGACY_TEST_ONLY );
		//let slide7 = pptx.addSlide('PLACEHOLDER_SLIDE');
	}
}
