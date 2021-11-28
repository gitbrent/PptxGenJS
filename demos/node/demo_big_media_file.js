import pptxgen from "pptxgenjs";

const VIDEO_FILES_PATH = "../common/media/earth-big.mp4"; // 17MB file, 1920x1080

const COVER_TEST =
	"data:image/png;base64,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";

console.log("Init...");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
let slide = pptx.addSlide();

slide.addText([{ text: VIDEO_FILES_PATH }], {
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
	path: VIDEO_FILES_PATH,
	cover: COVER_TEST,
});

slide.addMedia({
	x: 6.83,
	y: 2.0,
	w: 6,
	h: 3.38,
	type: "video",
	path: VIDEO_FILES_PATH,
	// cover: imgData.base64
});

console.log("Saving ppt 1...");
pptx.writeFile({ fileName: "big-file.pptx" });
