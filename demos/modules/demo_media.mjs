/**
 * NAME: demo_media.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.12.0
 * BLD.: 20230227
 */

import { IMAGE_PATHS, BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R, BASE_CODE_OPTS, BKGD_LTGRAY, COLOR_BLUE, CODE_STYLE, TITLE_STYLE } from "./enums.mjs";
import { COVER_AUDIO, COVER_VIDEO_16X9, COVER_YOUTUBE } from "./media.mjs";

export function genSlides_Media(pptx) {
	pptx.addSection({ title: "Media" });

	genSlide01(pptx);
	if (typeof window !== "undefined" && $ && $("#chkYoutube").prop("checked")) genSlide02(pptx);
	genSlide03(pptx);
	//if (window && window.location.href.indexOf("localhost:8000") > -1) genSlide03(pptx);
}

/**
 * SLIDE 1: Various Video Formats
 * @param {PptxGenJS} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Media" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html");
	slide.addTable([[{ text: "Media: Various Video Formats", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	slide.addText("Video: m4v", { x: 0.5, y: 0.6, w: 4.0, h: 0.4, color: "0088CC" });
	slide.addMedia({
		x: 0.5,
		y: 1.0,
		w: 4.0,
		h: 2.27,
		type: "video",
		path: IMAGE_PATHS.sample_m4v.path,
		cover: COVER_VIDEO_16X9,
	});

	slide.addText("Video: mpg", { x: 5.5, y: 0.6, w: 3.0, h: 0.4, color: "0088CC" });
	slide.addMedia({
		x: 5.5,
		y: 1.0,
		w: 3.0,
		h: 2.05,
		type: "video",
		path: IMAGE_PATHS.sample_mpg.path,
	});

	slide.addText("Video: mov", { x: 9.4, y: 0.6, w: 3.0, h: 0.4, color: "0088CC" });
	slide.addMedia({
		x: 9.4,
		y: 1.0,
		w: 3.0,
		h: 1.71,
		type: "video",
		path: IMAGE_PATHS.sample_mov.path,
	});

	slide.addText("Video: mp4", { x: 0.5, y: 3.6, w: 4.0, h: 0.4, color: "0088CC" });
	slide.addMedia({
		x: 0.5,
		y: 4.0,
		w: 4.0,
		h: 3.0,
		type: "video",
		path: IMAGE_PATHS.sample_mp4.path,
	});

	slide.addText("Video: avi", { x: 5.5, y: 3.6, w: 3.0, h: 0.4, color: "0088CC" });
	slide.addMedia({
		x: 5.5,
		y: 4.0,
		w: 3.0,
		h: 2.25,
		type: "video",
		path: IMAGE_PATHS.sample_avi.path,
	});
}

/**
 * SLIDE 2: YouTube
 * @param {PptxGenJS} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Media" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html");
	slide.addTable([[{ text: "Media: YouTube Embed", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	slide.addText("Online: YouTube", { ...{ x: 0.5, y: 0.75, h: 5.6, w: 12.3 }, ...TITLE_STYLE });
	// YouTube `link` is the embed URL (share > embed > copy URL like what you see below)
	slide.addMedia({ x: 2.1, y: 1.2, h: 5.1, w: 9.1, type: "online", link: "https://www.youtube.com/embed/g36-noRtKR4", cover: COVER_YOUTUBE });
	slide.addText(
		[{ text: 'slide.addMedia({ type: "online", link: "https://www.youtube.com/embed/g36-noRtKR4" })' }],
		{ ...BASE_CODE_OPTS, ...{ x: 0.5, y: 6.35, h: 0.4, w: 12.3 }, ...CODE_STYLE, ...{ align: 'center' } }
	);

	// FOOTER
	slide.addText("Note: YouTube videos require newer versions of PowerPoint (v16+/M365). Older versions will show content warning messages.", {
		shape: pptx.shapes.RECTANGLE,
		x: 0.0,
		y: 7.0,
		w: "100%",
		h: 0.53,
		color: "BF9000",
		fill: { color: "FFFCCC" },
		align: "center",
		fontSize: 12,
	});
}

/**
 * SLIDE 3: Various Audio Formats
 * @param {PptxGenJS} pptx
 */
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Media" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html");
	slide.addTable([[{ text: "Media: Various Audio Formats", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	slide.addText("Audio: mp3", { x: 0.5, y: 0.6, w: 4.0, h: 0.4, color: "0088CC" });
	slide.addMedia({
		x: 0.5,
		y: 1.0,
		w: 4.0,
		h: 4.0,
		type: "audio",
		path: IMAGE_PATHS.sample_mp3.path,
		cover: COVER_AUDIO,
	});

	slide.addText("Audio: wav", { x: 6.7, y: 0.6, w: 4.0, h: 0.4, color: "0088CC" });
	slide.addMedia({
		x: 6.7,
		y: 1.0,
		w: 4.0,
		h: 4.0,
		type: "audio",
		path: IMAGE_PATHS.sample_wav.path,
	});

	if (typeof window !== "undefined" && window.location.href.indexOf("gitbrent") > 0) {
		// TEST USING LOCAL FILES (OFFICE.COM)
		slide.addText('Audio: MP3 (path:"../media")', { x: 0.5, y: 4.6, w: 4.0, h: 0.4, color: "0088CC" });
		slide.addMedia({ x: 0.5, y: 5.0, w: 4.0, h: 0.3, type: "audio", path: "media/sample.mp3" });
	}
}

/**
 * SLIDE 3: Test large files are only added to export once
 * - filesize s/b ~24mb (the size of a single big-earth.mp4 file (17MB) plus other media files)
 * @param {PptxGenJS} pptx
 */
function genSlide_Test_LargeMedia(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Media" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html");
	slide.addTable([[{ text: "Media: Test: Large Files Only Added Once", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	slide.addText([{ text: IMAGE_PATHS.big_earth_mp4.path }], {
		x: 0.5,
		y: 0.5,
		w: 12.2,
		h: 1,
		fill: { color: "EEEEEE" },
		margin: 0,
		color: "000000",
	});

	slide.addMedia({
		x: 0.5,
		y: 2.0,
		w: 6,
		h: 3.38,
		type: "video",
		path: `${typeof window === "undefined" ? ".." : ""}${IMAGE_PATHS.big_earth_mp4.path}`, // NOTE: Node will throw exception when using "/" path
		cover: COVER_VIDEO_16X9,
	});

	slide.addMedia({
		x: 6.83,
		y: 2.0,
		w: 6,
		h: 3.38,
		type: "video",
		path: `${typeof window === "undefined" ? ".." : ""}${IMAGE_PATHS.big_earth_mp4.path}`, // NOTE: Node will throw exception when using "/" path
	});
}
