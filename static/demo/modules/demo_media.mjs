/**
 * NAME: demo_media.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.5.0
 * BLD.: 20210401
 */

import { IMAGE_PATHS, BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R } from "./enums.mjs";

export function genSlides_Media(pptx) {
	pptx.addSection({ title: "Media" });

	genSlide01(pptx);
	genSlide02(pptx);
}

/**
 * SLIDE 1: Video and YouTube
 * @param {PptxGenJS} pptx
 */
function genSlide01(pptx) {
	let slide1 = pptx.addSlide({ sectionTitle: "Media" });
	slide1.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html");
	slide1.addTable([[{ text: "Media: Misc Video Formats; YouTube", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	slide1.addText("Video: m4v", { x: 0.5, y: 0.6, w: 4.0, h: 0.4, color: "0088CC" });
	slide1.addMedia({
		x: 0.5,
		y: 1.0,
		w: 4.0,
		h: 2.27,
		type: "video",
		path: IMAGE_PATHS.sample_m4v.path,
	});

	slide1.addText("Video: mpg", { x: 5.5, y: 0.6, w: 3.0, h: 0.4, color: "0088CC" });
	slide1.addMedia({
		x: 5.5,
		y: 1.0,
		w: 3.0,
		h: 2.05,
		type: "video",
		path: IMAGE_PATHS.sample_mpg.path,
	});

	slide1.addText("Video: mov", { x: 9.4, y: 0.6, w: 3.0, h: 0.4, color: "0088CC" });
	slide1.addMedia({
		x: 9.4,
		y: 1.0,
		w: 3.0,
		h: 1.71,
		type: "video",
		path: IMAGE_PATHS.sample_mov.path,
	});

	slide1.addText("Video: mp4", { x: 0.5, y: 3.6, w: 4.0, h: 0.4, color: "0088CC" });
	slide1.addMedia({
		x: 0.5,
		y: 4.0,
		w: 4.0,
		h: 3.0,
		type: "video",
		path: IMAGE_PATHS.sample_mp4.path,
	});

	slide1.addText("Video: avi", { x: 5.5, y: 3.6, w: 3.0, h: 0.4, color: "0088CC" });
	slide1.addMedia({
		x: 5.5,
		y: 4.0,
		w: 3.0,
		h: 2.25,
		type: "video",
		path: IMAGE_PATHS.sample_avi.path,
	});

	// NOTE: Only generated on Node as I dont want everyone who downloads and runs this to be greated with an error!
	if (typeof window !== "undefined" && $ && $("#chkYoutube").prop("checked")) {
		slide1.addText("Online: YouTube", { x: 9.4, y: 3.6, w: 3.0, h: 0.4, color: "0088CC" });
		// Provide the usual options (locations and size), then pass the embed code from YouTube (it's on every video page)
		slide1.addMedia({ x: 9.4, y: 4.0, w: 3.0, h: 2.25, type: "online", link: "https://www.youtube.com/embed/Dph6ynRVyUc" });

		slide1.addText("**NOTE** YouTube videos will issue a content warning in older desktop PPT (they only work in PPT Online/Desktop v16+)", {
			shape: pptx.shapes.RECTANGLE,
			x: 0.0,
			y: 7.0,
			w: "100%",
			h: 0.53,
			fill: { color: "FFF000" },
			align: "center",
			fontSize: 12,
		});
	}
}

/**
 * SLIDE 2: Audio
 * @param {PptxGenJS} pptx
 */
function genSlide02(pptx) {
	let slide2 = pptx.addSlide({ sectionTitle: "Media" });
	slide2.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html");
	slide2.addTable([[{ text: "Media: Misc Audio Formats", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	slide2.addText("Audio: mp3", { x: 0.5, y: 0.6, w: 4.0, h: 0.4, color: "0088CC" });
	slide2.addMedia({
		x: 0.5,
		y: 1.0,
		w: 4.0,
		h: 0.3,
		type: "audio",
		path: IMAGE_PATHS.sample_mp3.path,
	});

	slide2.addText("Audio: wav", { x: 0.5, y: 2.6, w: 4.0, h: 0.4, color: "0088CC" });
	slide2.addMedia({
		x: 0.5,
		y: 3.0,
		w: 4.0,
		h: 0.3,
		type: "audio",
		path: IMAGE_PATHS.sample_wav.path,
	});

	if (typeof window !== "undefined" && window.location.href.indexOf("gitbrent") > 0) {
		// TEST USING LOCAL FILES (OFFICE.COM)
		slide2.addText('Audio: MP3 (path:"../media")', { x: 0.5, y: 4.6, w: 4.0, h: 0.4, color: "0088CC" });
		slide2.addMedia({ x: 0.5, y: 5.0, w: 4.0, h: 0.3, type: "audio", path: "media/sample.mp3" });
	}
}
