/**
 * NAME: demo_images.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.5.0
 * BLD.: 20210403
 */

/**
 * NOTES:
 * - Images can be pre-encoded into base64, so they do not have to be on the webserver etc. (saves generation time and resources!)
 * - This also has the benefit of being able to be any type (path:images can only be exported as PNG)
 * - Image source: either `data` or `path` is required
 */

import { IMAGE_PATHS, BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R, BASE_CODE_OPTS } from "./enums.mjs";
import { CHECKMARK_GRN, LOGO_STARLABS, SVG_BASE64, HYPERLINK_SVG } from "./media.mjs";

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

	slide.addTable([[{ text: "Image Examples: Misc Image Types", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", color: "0088CC" };

	// TOP: 1
	slide.addText("Type: Animated GIF", { x: 0.5, y: 0.6, w: 2.5, h: 0.4, color: "0088CC" });
	slide.addImage({ x: 1.0, y: 1.1, w: 1.5, h: 1.5, path: IMAGE_PATHS.gifAnimTrippy.path });
	slide.addText("(use slide Show)", {
		x: 1.0,
		y: 2.7,
		w: 1.5,
		h: 0.3,
		align: "center",
		fontSize: 10,
		color: "696969",
		fill: { color: "FFFCCC" },
	});

	// TOP: 2
	slide.addText("Type: GIF", { x: 4.35, y: 0.6, w: 1.4, h: 0.4, color: "0088CC" });
	slide.addImage({ x: 4.4, y: 1.05, w: 1.2, h: 1.2, path: IMAGE_PATHS.ccDjGif.path });

	// TOP: 3
	slide.addText("Type: base64 PNG", { x: 7.2, y: 0.6, w: 2.4, h: 0.4, color: "0088CC" });
	slide.addImage({ x: 7.87, y: 1.1, w: 1.0, h: 1.0, data: CHECKMARK_GRN });

	// TOP: 4
	slide.addText("Image Hyperlink", { x: 10.9, y: 0.6, w: 2.2, h: 0.4, color: "0088CC" });
	slide.addImage({
		x: 11.54,
		y: 1.2,
		w: 0.8,
		h: 0.8,
		data: HYPERLINK_SVG,
		hyperlink: { url: "https://github.com/gitbrent/pptxgenjs", tooltip: "Visit Homepage" },
	});

	// BOTTOM-LEFT:
	slide.addText("Type: JPG", { x: 0.5, y: 3.3, w: 4.5, h: 0.4, color: "0088CC" });
	slide.addImage({ path: IMAGE_PATHS.ccCopyRemix.path, x: 0.5, y: 3.8, w: 3.0, h: 3.07 });

	// BOTTOM-CENTER:
	slide.addText("Type: PNG", { x: 5.1, y: 3.3, w: 4.0, h: 0.4, color: "0088CC" });
	slide.addImage({ path: IMAGE_PATHS.wikimedia1.path, x: 5.1, y: 3.8, w: 3.0, h: 2.78 });

	// BOTTOM-RIGHT:
	slide.addText("Type: SVG", { x: 9.5, y: 3.3, w: 4.0, h: 0.4, color: "0088CC" });
	slide.addImage({ path: IMAGE_PATHS.wikimedia_svg.path, x: 9.5, y: 3.8, w: 2.0, h: 2.0 }); // TEST: `path`
	slide.addImage({ data: SVG_BASE64, x: 11.1, y: 5.1, w: 1.5, h: 1.5 }); // TEST: `data`

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

	slide.addTable([[{ text: "Image Examples: Image Sizing/Rounding", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", w: 1, h: 1, color: "0088CC" };

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
	slide.addImage({ path: IMAGE_PATHS.ccLogo.path, x: 9.9, y: 1.1, w: 2.5, h: 2.5, rounding: true });
}

/**
 * SLIDE 3: Image Rotation
 * @param {PptxGenJS} pptx
 */
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });

	slide.addTable([[{ text: "Image Examples: Image Rotation", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", w: 1, h: 1, color: "0088CC" };

	// EXAMPLES
	slide.addText("Rotate: `rotate:45`, `rotate:180`, `rotate:315`", { x: 0.5, y: 0.6, w: 6.0, h: 0.3, color: "0088CC" });
	slide.addImage({ path: IMAGE_PATHS.tokyoSubway.path, x: 0.78, y: 2.46, w: 4.3, h: 3, rotate: 45 });
	slide.addImage({ path: IMAGE_PATHS.tokyoSubway.path, x: 4.52, y: 2.25, w: 4.3, h: 3, rotate: 180 });
	slide.addImage({ path: IMAGE_PATHS.tokyoSubway.path, x: 8.25, y: 2.84, w: 4.3, h: 3, rotate: 315 });
}

/**
 * SLIDE 4: Image URLs
 * @param {PptxGenJS} pptx
 */
function genSlide04(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Images" });

	slide.addTable([[{ text: "Image Examples: Image URLs", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html");
	slide.slideNumber = { x: "50%", y: "95%", color: "0088CC" };

	// TOP-LEFT: jpg
	slide.addText([{ text: `path:"${IMAGE_PATHS.ccLogo.path}"` }], { ...BASE_CODE_OPTS, ...{ x: 0.5, y: 0.6, w: 3.3, h: 0.7, fontSize: 11 } });
	slide.addImage({ path: IMAGE_PATHS.ccLogo.path, x: 0.5, y: 1.44, h: 2.5, w: 3.33 });

	// TOP-CENTER: png
	slide.addText([{ text: `path:"${IMAGE_PATHS.wikimedia2.path}"` }], { ...BASE_CODE_OPTS, ...{ x: 4.55, y: 0.6, w: 3.27, h: 0.7, fontSize: 11 } });
	slide.addImage({ path: IMAGE_PATHS.wikimedia2.path, x: 4.55, y: 1.44, h: 2.5, w: 3.27 });

	// TOP-RIGHT: relative-path test
	slide.addText([{ text: `path:"${IMAGE_PATHS.ccLicenseComp.path}"` }], {
		...BASE_CODE_OPTS,
		...{ x: 8.55, y: 0.6, w: 4.28, h: 0.7, fontSize: 11 },
	});
	// NOTE: Node will throw exception when using "/" path
	slide.addImage({ path: `${typeof window === "undefined" ? ".." : ""}${IMAGE_PATHS.ccLicenseComp.path}`, x: 8.55, y: 1.43, h: 2.51, w: 4.28 });

	// BOTTOM: wide, url-sourced
	slide.addText(
		[
			{ text: '// Test: URL variables, plus more than one ".jpg"', options: { breakLine: true } },
			{ text: `path:"${IMAGE_PATHS.sydneyBridge.path}"` },
		],
		{
			...BASE_CODE_OPTS,
			...{ x: 0.5, y: 4.2, w: 12.33, h: 0.8, fontSize: 11 },
		}
	);
	slide.addImage({ path: IMAGE_PATHS.sydneyBridge.path, x: 0.5, y: 5.16, h: 1.8, w: 12.33 });
}
