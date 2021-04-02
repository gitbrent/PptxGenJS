/**
 * NAME: demo_images.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.5.0
 * BLD.: 20210401
 */

/**
 * NOTES:
 * - Images can be pre-encoded into base64, so they do not have to be on the webserver etc. (saves generation time and resources!)
 * - This also has the benefit of being able to be any type (path:images can only be exported as PNG)
 * - Image source: either `data` or `path` is required
 */

import { gPaths, gOptsTabOpts, gOptsTextL, gOptsTextR, gOptsCode } from "./enums.js";
import { checkGreen, LOGO_STARLABS, svgBase64, svgHyperlinkImage } from "./media.js";

export function genSlides_Image(pptx) {
	pptx.addSection({ title: "Images" });

	genSlide01(pptx);
	genSlide02(pptx);
	genSlide03(pptx);
	genSlide04(pptx);
}

/**
 * SLIDE 1:
 * @param {PptxGenJS} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", color: "0088CC" };
	slide.addTable([[{ text: "Image Examples: Misc Image Types", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

	// TOP: 1
	slide.addText("Type: Animated GIF", { x: 0.5, y: 0.6, w: 2.5, h: 0.4, color: "0088CC" });
	slide.addImage({ x: 1.0, y: 1.1, w: 1.5, h: 1.5, path: gPaths.gifAnimTrippy.path });
	slide.addText("(use slide Show)", {
		x: 1.0,
		y: 2.7,
		w: 1.5,
		h: 0.3,
		color: "696969",
		fill: { color: "FFFCCC" },
		align: "center",
		fontSize: 10,
	});

	// TOP: 2
	slide.addText("Type: GIF", { x: 4.35, y: 0.6, w: 1.4, h: 0.4, color: "0088CC" });
	slide.addImage({ x: 4.4, y: 1.05, w: 1.2, h: 1.2, path: gPaths.ccDjGif.path });

	// TOP: 3
	slide.addText("Type: base64 PNG", { x: 7.2, y: 0.6, w: 2.4, h: 0.4, color: "0088CC" });
	slide.addImage({ x: 7.87, y: 1.1, w: 1.0, h: 1.0, data: checkGreen });

	// TOP: 4
	slide.addText("Image Hyperlink", { x: 10.9, y: 0.6, w: 2.2, h: 0.4, color: "0088CC" });
	slide.addImage({
		x: 11.54,
		y: 1.2,
		w: 0.8,
		h: 0.8,
		data: svgHyperlinkImage,
		hyperlink: { url: "https://github.com/gitbrent/pptxgenjs", tooltip: "Visit Homepage" },
	});

	// BOTTOM-LEFT:
	slide.addText("Type: JPG", { x: 0.5, y: 3.3, w: 4.5, h: 0.4, color: "0088CC" });
	slide.addImage({ path: gPaths.ccCopyRemix.path, x: 0.5, y: 3.8, w: 3.0, h: 3.07 });

	// BOTTOM-CENTER:
	slide.addText("Type: PNG", { x: 5.1, y: 3.3, w: 4.0, h: 0.4, color: "0088CC" });
	slide.addImage({ path: gPaths.wikimedia1.path, x: 5.1, y: 3.8, w: 3.0, h: 2.78 });

	// BOTTOM-RIGHT:
	slide.addText("Type: SVG", { x: 9.5, y: 3.3, w: 4.0, h: 0.4, color: "0088CC" });
	slide.addImage({ path: gPaths.wikimedia_svg.path, x: 9.5, y: 3.8, w: 2.0, h: 2.0 }); // TEST: `path`
	slide.addImage({ data: svgBase64, x: 11.1, y: 5.1, w: 1.5, h: 1.5 }); // TEST: `data`

	// TEST: Ensure framework corrects for missing all header
	// (Please **DO NOT** pass base64 data without the header! This is a JUNK TEST!)
	//slide.addImage({ x:5.2, y:2.6, w:0.8, h:0.8, data:'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAAjcAAAI3AGf6F88AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAANVQTFRF////JLaSIJ+AIKqKKa2FKLCIJq+IJa6HJa6JJa6IJa6IJa2IJa6IJa6IJa6IJa6IJa6IJa6IJq6IKK+JKK+KKrCLLrGNL7KOMrOPNrSRN7WSPLeVQrmYRLmZSrycTr2eUb6gUb+gWsKlY8Wqbsmwb8mwdcy0d8y1e863g9G7hdK8htK9i9TAjNTAjtXBktfEntvKoNzLquDRruHTtePWt+TYv+fcx+rhyOvh0e7m1e/o2fHq4PTu5PXx5vbx7Pj18fr49fv59/z7+Pz7+f38/P79/f7+dNHCUgAAABF0Uk5TAAcIGBktSYSXmMHI2uPy8/XVqDFbAAABB0lEQVQ4y42T13qDMAyFZUKMbebp3mmbrnTvlY60TXn/R+oFGAyYzz1Xx/wylmWJqBLjUkVpGinJGXXliwSVEuG3sBdkaCgLPJMPQnQUDmo+jGFRPKz2WzkQl//wQvQoLPII0KuAiMjP+gMyn4iEFU1eAQCCiCU2fpCfFBVjxG18f35VOk7Swndmt9pKUl2++fG4qL2iqMPXpi8r1SKitDDne/rT8vPbRh2d6oC7n6PCLNx/bsEM0Edc5DdLAHD9tWueF9VJjmdP68DZ77iRkDKuuT19Hx3mx82MpVmo1Yfv+WXrSrxZ6slpiyes77FKif88t7Nh3C3nbFp327sHxz167uHtH/8/eds7gGsUQbkAAAAASUVORK5CYII=' });
	// NEGATIVE-TEST:
	//slide.addImage({ data:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/images/doh_this_isnt_base64_data.gif',  x:0.5, y:0.5, w:1.0, h:1.0 });
}

/**
 * SLIDE 2: Image Sizing
 * @param {PptxGenJS} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", w: 1, h: 1, color: "0088CC" };
	slide.addTable([[{ text: "Image Examples: Image Sizing/Rounding", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

	// TOP: 1
	slide.addText("Sizing: Orig `w:6, h:2.7`", { x: 0.5, y: 0.6, w: 3.0, h: 0.3, color: "0088CC" });
	slide.addImage({ data: LOGO_STARLABS, x: 0.5, y: 1.1, w: 6.0, h: 2.69 });

	// TOP: 2
	slide.addText("Sizing: `contain, w:3`", { x: 0.6, y: 4.25, w: 3.0, h: 0.3, color: "0088CC" });
	slide.addShape(pptx.shapes.RECTANGLE, { x: 0.6, y: 4.65, w: 3, h: 2, fill: { color: "F1F1F1" } });
	slide.addImage({ data: LOGO_STARLABS, x: 0.6, y: 4.65, w: 5.0, h: 1.5, sizing: { type: "contain", w: 3, h: 2 } });

	// TOP: 3
	slide.addText("Sizing: `cover, w:3, h:2`", { x: 5.3, y: 4.25, w: 3.0, h: 0.3, color: "0088CC" });
	slide.addShape(pptx.shapes.RECTANGLE, { x: 5.3, y: 4.65, w: 3, h: 2, fill: { color: "F1F1F1" } });
	slide.addImage({ data: LOGO_STARLABS, x: 5.3, y: 4.65, w: 3.0, h: 1.5, sizing: { type: "cover", w: 3, h: 2 } });

	// TOP: 4
	slide.addText("Sizing: `crop, w:3, h:2`", { x: 10.0, y: 4.25, w: 3.0, h: 0.3, color: "0088CC" });
	slide.addShape(pptx.shapes.RECTANGLE, { x: 10, y: 4.65, w: 3, h: 1.5, fill: { color: "F1F1F1" } });
	slide.addImage({ data: LOGO_STARLABS, x: 10.0, y: 4.65, w: 5.0, h: 1.5, sizing: { type: "crop", w: 3, h: 1.5, x: 0.5, y: 0.5 } });

	// TOP-RIGHT:
	slide.addText("Rounding: `rounding:true`", { x: 10.0, y: 0.6, w: 3.0, h: 0.3, color: "0088CC" });
	slide.addImage({ path: gPaths.ccLogo.path, x: 9.9, y: 1.1, w: 2.5, h: 2.5, rounding: true });
}

/**
 * SLIDE 3: Image Rotation
 * @param {PptxGenJS} pptx
 */
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", w: 1, h: 1, color: "0088CC" };
	slide.addTable([[{ text: "Image Examples: Image Rotation", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

	// EXAMPLES
	slide.addText("Rotate: `rotate:45`, `rotate:180`, `rotate:315`", { x: 0.5, y: 0.6, w: 6.0, h: 0.3, color: "0088CC" });
	slide.addImage({ path: gPaths.tokyoSubway.path, x: 0.78, y: 2.46, w: 4.3, h: 3, rotate: 45 });
	slide.addImage({ path: gPaths.tokyoSubway.path, x: 4.52, y: 2.25, w: 4.3, h: 3, rotate: 180 });
	slide.addImage({ path: gPaths.tokyoSubway.path, x: 8.25, y: 2.84, w: 4.3, h: 3, rotate: 315 });
}

/**
 * SLIDE 4: Image URLs
 * @param {PptxGenJS} pptx
 */
function genSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", color: "0088CC" };
	slide.addTable([[{ text: "Image Examples: Image URLs", options: gOptsTextL }, gOptsTextR]], gOptsTabOpts);

	// TOP-LEFT:
	let objCodeEx1 = { x: 0.5, y: 0.6, w: 6.0, h: 0.6 };
	Object.keys(gOptsCode).forEach((key) => {
		objCodeEx1[key] = gOptsCode[key];
	});
	slide.addText('path:"' + gPaths.ccLogo.path + '"', objCodeEx1);
	slide.addImage({ path: gPaths.ccLogo.path, x: 1.84, y: 1.3, h: 2.5, w: 3.33 });

	// TOP-RIGHT:
	let objCodeEx2 = { x: 6.9, y: 0.6, w: 5.93, h: 0.6 };
	Object.keys(gOptsCode).forEach((key) => {
		objCodeEx2[key] = gOptsCode[key];
	});
	slide.addText('path:"' + gPaths.wikimedia2.path + '"', objCodeEx2);
	slide.addImage({ path: gPaths.wikimedia2.path, x: 8.23, y: 1.3, h: 2.5, w: 3.27 });

	// BTM-LEFT:
	let objCodeEx3 = { x: 0.5, y: 4.2, w: 12.33, h: 0.8 };
	Object.keys(gOptsCode).forEach((key) => {
		objCodeEx3[key] = gOptsCode[key];
	});
	slide.addText('// Test: URL variables, plus more than one ".jpg"\npath:"' + gPaths.sydneyBridge.path + '"', objCodeEx3);
	slide.addImage({ path: gPaths.sydneyBridge.path, x: 0.5, y: 5.1, h: 1.8, w: 12.33 });

	// BOTTOM-CENTER:
	if (typeof window !== "undefined" && window.location.href.indexOf("gitbrent") > 0) {
		// TEST USING RELATIVE PATHS/LOCAL FILES (OFFICE.COM)
		slide.addText('Type: PNG (path:"../images")', { x: 6.6, y: 2.7, w: 4.5, h: 0.4, color: "CC0033" });
		slide.addImage({ path: gPaths.ccLicenseComp.path, x: 6.6, y: 3.2, w: 6.3, h: 3.7 });
	}
}
