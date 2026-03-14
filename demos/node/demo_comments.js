import pptxgen from "../../dist/pptxgen.es.js";

const pptx = new pptxgen();
const exportName = "PptxGenJS_Demo_Comments";

pptx.author = "PptxGenJS Comment Demo";
pptx.company = "PptxGenJS";

const slide = pptx.addSlide();

slide.addText("Element Comments Demo", {
	x: 0.5,
	y: 0.3,
	w: 6.5,
	h: 0.5,
	fontSize: 24,
	bold: true,
	color: "1F1F1F",
	comment: {
		text: "Opening title could be shortened for the executive version.",
		authorName: "Design Review",
		authorInitials: "DR",
	},
});

slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
	x: 0.6,
	y: 1.1,
	w: 4.2,
	h: 1.0,
	fill: { color: "E8F2FF" },
	line: { color: "5B9BD5", width: 1 },
	comment: "This summary box needs final compliance wording.",
});

slide.addText("Summary callout", {
	x: 0.9,
	y: 1.38,
	w: 3.4,
	h: 0.3,
	fontSize: 18,
	bold: true,
	color: "1F1F1F",
});

slide.addTable(
	[
		[{ text: "Segment", options: { bold: true, fill: { color: "D9EAD3" } } }, { text: "Value", options: { bold: true, fill: { color: "D9EAD3" } } }],
		["Commercial", "$4.2M"],
		["Private Banking", "$2.8M"],
		["Treasury", "$1.1M"],
	],
	{
		x: 0.7,
		y: 2.5,
		w: 3.8,
		border: { type: "solid", color: "8FAADC", pt: 1 },
		fontSize: 14,
		comment: {
			text: "Reconcile these figures against the month-end close before release.",
			authorName: "Finance QA",
			authorInitials: "FQ",
		},
	}
);

slide.addImage({
	path: "../common/images/cc_logo.jpg",
	x: 5.6,
	y: 1.4,
	w: 3.3,
	h: 2.0,
	comment: "Replace with the approved 2026 partner logo lockup.",
});

slide.addText("Each commented element should show a native PowerPoint comment marker near its top-right corner.", {
	x: 0.7,
	y: 5.35,
	w: 8.8,
	h: 0.6,
	fontSize: 12,
	color: "555555",
});

pptx.writeFile({ fileName: exportName })
	.then((fileName) => {
		console.log(`COMMENTS DEMO exported: ${fileName}`);
	})
	.catch((err) => {
		console.error(`ERROR: ${err}`);
		process.exitCode = 1;
	});