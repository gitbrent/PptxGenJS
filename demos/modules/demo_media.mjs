/**
 * NAME: demo_media.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.12.0
 * BLD.: 20230314
 */

/**
 * PowerPoint supports a variety of video formats.
 * The supported video file formats can depend on the version of PowerPoint being used.
 *
 * Here are some of the most common video file formats that are supported by PowerPoint:
 * .avi (audio video interleave)
 * .mp4 (MPEG-4 video)
 * .mov (QuickTime movie)
 * .mv4 (iTunes movie, Apple's version on MP4)
 * .wmv (windows media video)	[[NOT USED IN DEMO]]
 * .mpg or .mpeg (DVD video)	[[NOT USED IN DEMO]]
 *
 * It's worth noting that even if a video file format is supported by PowerPoint,
 * you may still encounter issues with playing the video if the video is encoded using a codec that is not supported by the computer you are using to present the slideshow.
 * It's a good idea to test your slideshow on the computer you will be using to present it to ensure that your videos will play correctly.
 */

/**
 * PowerPoint supports several audio file formats, including:
 * - MP3  (MPEG Audio Layer III)
 * - WAV  (Waveform Audio Format) WAV files can contain a variety of audio codecs, including PCM, ADPCM, and others. They are widely supported by media players and software applications on both Windows and Mac operating systems.
 * - AIFF (Audio Interchange File Format) AIFF files can be played on both Mac and Windows computers, as well as on many other types of devices. They are often used in music production and editing applications, as well as for storing high-quality audio recordings.
 * - MIDI (Musical Instrument Digital Interface) MIDI files are typically small in size compared to other audio formats, and they can be edited and manipulated using specialized software. They are often used in music production and composition, as well as in live performances
 * - WMA  (Windows Media Audio) [[not demoed]]
 * In addition to these formats, PowerPoint also supports embedding audio from online sources like YouTube and SoundCloud,
 * as well as recording audio directly within the presentation using the built-in audio recording feature.
 */

import { IMAGE_PATHS, BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R, BASE_CODE_OPTS, CODE_STYLE, TITLE_STYLE } from "./enums.mjs";
import { COVER_AUDIO, COVER_AUDIO_ROUND, COVER_VIDEO_16X9, COVER_VIDEO_MP4, COVER_YOUTUBE } from "./media.mjs";

export function genSlides_Media(pptx) {
	pptx.addSection({ title: "Media" });

	genSlide01(pptx);
	genSlide02(pptx);
	if (typeof window !== "undefined" && $ && $("#chkYoutube").prop("checked")) genSlide03(pptx);
	//if (window && window.location.href.indexOf("localhost:8000") > -1) genSlide03(pptx);
}

/**
 * SLIDE 1: Various Video Formats
 * @param {PptxGenJS} pptx
 */
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Media" });
	slide.addTable([[{ text: "Media Examples: Video Types", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html\r\nIt's worth noting that even if a video file format is supported by PowerPoint, you may still encounter issues with playing the video if the video is encoded using a codec that is not supported by the computer you are using to present the slideshow. It's a good idea to test your slideshow on the computer you will be using to present it to ensure that your videos will play correctly.");

	slide.addText([{ text: "Type: m4v" }], { ...BASE_CODE_OPTS, ...{ x: 0.5, y: 0.6, h: 0.4, w: 3.56 }, ...TITLE_STYLE });
	slide.addMedia({ x: 0.5, y: 1.0, h: 2.0, w: 3.56, type: "video", path: IMAGE_PATHS.sample_m4v.path, cover: COVER_VIDEO_16X9 });
	slide.addText([{ text: "`cover` image provided" }], { ...BASE_CODE_OPTS, ...{ x: 0.5, y: 3.0, h: 0.4, w: 3.56 }, ...CODE_STYLE });

	slide.addText([{ text: "Type: m4v" }], { ...BASE_CODE_OPTS, ...{ x: 9.3, y: 0.6, h: 0.4, w: 3.56 }, ...TITLE_STYLE });
	slide.addMedia({ x: 9.3, y: 1.0, h: 2.0, w: 3.56, type: "video", path: IMAGE_PATHS.sample_m4v.path });
	slide.addText([{ text: "no `cover` image provided" }], { ...BASE_CODE_OPTS, ...{ x: 9.3, y: 3.0, h: 0.4, w: 3.56 }, ...CODE_STYLE });

	// BOTTOM-ROW

	slide.addText([{ text: "Type: mp4" }], { ...BASE_CODE_OPTS, ...{ x: 0.5, y: 3.85, h: 0.4, w: 3.6 }, ...TITLE_STYLE });
	slide.addMedia({
		x: 0.5,
		y: 4.25,
		h: 2.7,
		w: 3.6,
		type: "video",
		path: IMAGE_PATHS.sample_mp4.path,
		cover: COVER_VIDEO_MP4,
	});

	slide.addText([{ text: "Type: avi" }], { ...BASE_CODE_OPTS, ...{ x: 4.79, y: 3.85, h: 0.4, w: 3.6 }, ...TITLE_STYLE });
	slide.addMedia({
		x: 4.79,
		y: 4.25,
		h: 2.7,
		w: 3.6,
		type: "video",
		path: IMAGE_PATHS.sample_avi.path,
	});

	slide.addText([{ text: "Type: mov" }], { ...BASE_CODE_OPTS, ...{ x: 9.08, y: 3.85, h: 0.4, w: 3.75 }, ...TITLE_STYLE });
	slide.addMedia({
		x: 9.08,
		y: 4.25,
		h: 2.7,
		w: 3.75,
		type: "video",
		path: IMAGE_PATHS.sample_mov.path,
	});
}

/**
 * SLIDE 2: Various Audio Typrs
 * @param {PptxGenJS} pptx
 */
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Media" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html");
	slide.addTable([[{ text: "Media Examples: Audio Types", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	slide.addText([{ text: "Type: mp3" }], { ...BASE_CODE_OPTS, ...{ x: 0.5, y: 0.6, h: 0.4, w: 3.5 }, ...TITLE_STYLE });
	slide.addMedia({
		x: 0.5,
		y: 1.0,
		h: 3.5,
		w: 3.5,
		type: "audio",
		path: IMAGE_PATHS.sample_mp3.path,
		cover: COVER_AUDIO,
	});

	slide.addText([{ text: "Type: aiff" }], { ...BASE_CODE_OPTS, ...{ x: 4.92, y: 0.6, h: 3.9, w: 3.5 }, ...TITLE_STYLE });
	slide.addMedia({
		x: 4.92,
		y: 1.0,
		h: 3.5,
		w: 3.5,
		type: "audio",
		path: IMAGE_PATHS.sample_aif.path,
		cover: COVER_AUDIO_ROUND,
	});

	slide.addText([{ text: "Type: wav" }], { ...BASE_CODE_OPTS, ...{ x: 9.33, y: 0.6, h: 0.4, w: 3.5 }, ...TITLE_STYLE });
	slide.addMedia({
		x: 9.33,
		y: 1.0,
		h: 3.5,
		w: 3.5,
		type: "audio",
		path: IMAGE_PATHS.sample_wav.path,
		cover: COVER_AUDIO,
	});

	if (typeof window !== "undefined" && window.location.href.indexOf("gitbrent") > 0) {
		// TEST USING LOCAL FILES (OFFICE.COM)
		slide.addText('Audio: MP3 (path:"../media")', { x: 0.5, y: 4.6, w: 4.0, h: 0.4, color: "0088CC" });
		slide.addMedia({ x: 0.5, y: 5.0, w: 4.0, h: 0.3, type: "audio", path: "media/sample.mp3" });
	}
}

/**
 * SLIDE 3: YouTube
 * @param {PptxGenJS} pptx
 */
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Media" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html");
	slide.addTable([[{ text: "Media Examples: YouTube Embed", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

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
