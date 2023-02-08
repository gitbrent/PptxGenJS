/**
 * NAME: masters.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.6.0
 * BLD.: 20210421
 */

import { IMAGE_PATHS } from "../modules/enums.mjs";
import { STARLABS_LOGO_SM } from "../modules/media.mjs";

export function createMasterSlides(pptx) {
	let objBkg = { path: IMAGE_PATHS.starlabsBkgd.path };
	let objImg = { path: IMAGE_PATHS.starlabsLogo.path, x: 4.6, y: 3.5, w: 4, h: 1.8 };

	// TITLE_SLIDE
	pptx.defineSlideMaster({
		title: "TITLE_SLIDE",
		background: objBkg,
		//bkgd: objBkg, // TEST: @deprecated
		objects: [
			//{ 'line':  { x:3.5, y:1.0, w:6.0, h:0.0, line:{color:'0088CC'}, lineSize:5 } },
			//{ 'chart': { type:'PIE', data:[{labels:['R','G','B'], values:[10,10,5]}], options:{x:11.3, y:0.0, w:2, h:2, dataLabelFontSize:9} } },
			//{ 'image': { x:11.3, y:6.4, w:1.67, h:0.75, data:STARLABS_LOGO_SM } },
			{ rect: { x: 0.0, y: 5.7, w: "100%", h: 0.75, fill: { color: "F1F1F1" } } },
			{
				text: {
					text: "Global IT & Services :: Status Report",
					options: {
						x: 0.0,
						y: 5.7,
						w: "100%",
						h: 0.75,
						fontFace: "Arial",
						color: "363636",
						fontSize: 20,
						align: "center",
						valign: "middle",
						margin: 0,
					},
				},
			},
		],
	});

	// MASTER_PLAIN
	pptx.defineSlideMaster({
		title: "MASTER_PLAIN",
		background: { fill: "F1F1F1" }, // [[BACKWARDS-COMPAT/DEPRECATED CHECK:changed to `color` in v3.5.0]]
		margin: [0.5, 0.25, 1.0, 0.25],
		objects: [
			{ rect: { x: 0.0, y: 6.9, w: "100%", h: 0.6, fill: { color: "003b75" } } },
			{ image: { x: 11.45, y: 5.95, w: 1.67, h: 0.75, data: STARLABS_LOGO_SM } },
			{
				text: {
					options: { x: 0, y: 6.9, w: "100%", h: 0.6, align: "center", valign: "middle", color: "FFFFFF", fontSize: 12 },
					text: "S.T.A.R. Laboratories - Confidential",
				},
			},
		],
		slideNumber: { x: 0.6, y: 7.1, color: "FFFFFF", fontFace: "Arial", fontSize: 10, align: "center" },
	});

	// MASTER_AUTO_PAGE_TABLE_PLACEHOLDER
	pptx.defineSlideMaster({
		title: "MASTER_AUTO_PAGE_TABLE_PLACEHOLDER",
		background: { fill: "F1F1F1" },
		margin: [0.5, 0.25, 1.0, 0.25],
		objects: [
			{ rect: { x: 0.0, y: 6.9, w: "100%", h: 0.6, fill: { color: "003b75" } } },
			{ image: { x: 11.45, y: 5.95, w: 1.67, h: 0.75, data: STARLABS_LOGO_SM } },
			{
				placeholder: {
					options: { name: "footer", x: 0, y: 6.9, w: "100%", h: 0.6, align: "center", valign: "middle", color: "FFFFFF", fontSize: 12 },
					text: "(footer placeholder)",
				},
			},
		],
		slideNumber: { x: 0.6, y: 7.1, color: "FFFFFF", fontFace: "Arial", fontSize: 10, align: "center" },
	});

	// MASTER_SLIDE (MASTER_PLACEHOLDER)
	pptx.defineSlideMaster({
		title: "MASTER_SLIDE",
		background: { color: "E1E1E1", transparency: 50 },
		margin: [0.5, 0.25, 1.0, 0.25],
		slideNumber: { x: 0.6, y: 7.1, color: "FFFFFF", fontFace: "Arial", fontSize: 10, bold: true },
		objects: [
			//{ 'image': { x:11.45, y:5.95, w:1.67, h:0.75, data:STARLABS_LOGO_SM } },
			{
				rect: { x: 0.0, y: 6.9, w: "100%", h: 0.6, fill: { color: "003b75" } },
			},
			{
				text: {
					options: { x: 0, y: 6.9, w: "100%", h: 0.6, align: "center", valign: "middle", color: "FFFFFF", fontSize: 12 },
					text: "S.T.A.R. Laboratories - Confidential",
				},
			},
			{
				placeholder: {
					options: {
						name: "header",
						type: "title",
						x: 0.6,
						y: 0.2,
						w: 12,
						h: 1.0,
						margin: 0,
						align: "center",
						valign: "middle",
						color: "404040",
						//fontSize: 18,
					},
					text: "", // USAGE: Leave blank to have powerpoint substitute default placeholder text (ex: "Click to add title")
				},
			},
			{
				placeholder: {
					options: { name: "body", type: "body", x: 0.6, y: 1.5, w: 12, h: 5.25, fontSize: 28 },
					text: "(supports custom placeholder text!)",
				},
			},
		],
	});

	// THANKS_SLIDE (THANKS_PLACEHOLDER)
	pptx.defineSlideMaster({
		title: "THANKS_SLIDE",
		background: { color: "36ABFF" }, // CORRECT WAY TO SET BACKGROUND COLOR
		//bkgd: "36ABFF", // [[BACKWARDS-COMPAT/DEPRECATED/UAT (`bkgd` will be removed in v4.x)]] **DO NOT USE THIS IN YOUR CODE**
		objects: [
			{ rect: { x: 0.0, y: 3.4, w: "100%", h: 2.0, fill: { color: "FFFFFF" } } },
			{ image: objImg },
			{
				placeholder: {
					options: {
						name: "thanksText",
						type: "title",
						x: 0.0,
						y: 0.9,
						w: "100%",
						h: 1,
						fontFace: "Arial",
						color: "FFFFFF",
						fontSize: 60,
						align: "center",
					},
				},
			},
			{
				placeholder: {
					options: {
						name: "body",
						type: "body",
						x: 0.0,
						y: 6.45,
						w: "100%",
						h: 1,
						fontFace: "Courier",
						color: "FFFFFF",
						fontSize: 32,
						align: "center",
					},
					text: "(add homepage URL)",
				},
			},
		],
	});

	// MARGIN_SLIDE (used for demo/test)
	const MARGINS = [0.5, 0.5, 0.5, 0.5];
	const TEXT_PROPS = {
		shape: pptx.shapes.RECTANGLE,
		fill: { color: "FFFCCC" },
		color: "9f9f9f",
		align: "center",
		fontFace: "Courier New",
		fontSize: 10,
	};
	pptx.defineSlideMaster({
		title: "MARGIN_SLIDE",
		background: { color: "FFFFFF" },
		margin: MARGINS,
		objects: [
			{ text: { text: "(margin-top)", options: { ...TEXT_PROPS, ...{ x: 0, y: 0, w: "100%", h: MARGINS[0] } } } },
			{ text: { text: "(margin-btm)", options: { ...TEXT_PROPS, ...{ x: 0, y: 7.5 - MARGINS[2], w: "100%", h: MARGINS[2], flipV: true } } } },
		],
	});

	// MARGIN_SLIDE_STARTY15 (used for demo/test)
	pptx.defineSlideMaster({
		title: "MARGIN_SLIDE_STARTY15",
		background: { color: "FFFFFF" },
		margin: MARGINS,
		objects: [
			{ text: { text: "(4.0 inches H)", options: { ...TEXT_PROPS, ...{ x: 0, y: 0, w: 1, h: 4.0 } } } },
			{ text: { text: "(1.5 inches H)", options: { ...TEXT_PROPS, ...{ x: 1, y: 0, w: 1, h: 1.5 } } } },
			{ text: { text: "(margin-top)", options: { ...TEXT_PROPS, ...{ x: 0, y: 0, w: "100%", h: MARGINS[0] } } } },
			{ text: { text: "(margin-btm)", options: { ...TEXT_PROPS, ...{ x: 0, y: 7.5 - MARGINS[2], w: "100%", h: MARGINS[2], flipV: true } } } },
		],
	});

	// PLACEHOLDER_SLIDE
	/* FUTURE: ISSUE#599
		pptx.defineSlideMaster({
		  title : 'PLACEHOLDER_SLIDE',
		  margin: [0.5, 0.25, 1.00, 0.25],
		  bkgd  : 'FFFFFF',
		  objects: [
			  { 'placeholder':
			  	{
					options: {type:'body'},
					image: {x:11.45, y:5.95, w:1.67, h:0.75, data:STARLABS_LOGO_SM}
				}
			},
			  { 'placeholder':
				  {
					  options: { name:'body', type:'body', x:0.6, y:1.5, w:12, h:5.25 },
					  text: '(supports custom placeholder text!)'
				  }
			  }
		  ],
		  slideNumber: { x:1.0, y:7.0, color:'FFFFFF' }
	  });*/

	// MISC: Only used for Issues, ad-hoc slides etc (for screencaps)
	pptx.defineSlideMaster({
		title: "DEMO_SLIDE",
		objects: [
			{ rect: { x: 0.0, y: 7.1, w: "100%", h: 0.4, fill: { color: "F1F1F1" } } },
			{
				text: {
					text: "PptxGenJS - JavaScript PowerPoint Library - (github.com/gitbrent/PptxGenJS)",
					options: { x: 0.0, y: 7.1, w: "100%", h: 0.4, color: "6c6c6c", fontSize: 10, align: "center" },
				},
			},
		],
	});
}

/**
 * Test Slide BACKGROUNDS
 */
export function testSlideBackgrounds(pptx) {
	let slide1 = pptx.addSlide();
	slide1.bkgd = "909090";
	slide1.addText([{ text: "TEST `bkgd:string`" }], { x: 1, y: 1, w: "80%", h: 3, align: "center", fill: { color: "a1a1a1" } });

	let slide2 = pptx.addSlide();
	slide2.background = { fill: "909090" };
	slide2.addText([{ text: "TEST `background.fill`" }], { x: 1, y: 1, w: "80%", h: 3, align: "center", fill: { color: "a1a1a1" } });

	let slide3 = pptx.addSlide();
	slide3.background = { color: "909090", transparency: 50 };
	slide3.addText([{ text: "TEST `background`[correct]" }], { x: 1, y: 1, w: "80%", h: 3, align: "center", fill: { color: "a1a1a1" } });
}
