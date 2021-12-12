/**
 * browser.js
 * module for /demo/browser/index.html
 */
import { execGenSlidesFuncs, runEveryTest } from "../../modules/demos.mjs";
import { TABLE_NAMES_F, TABLE_NAMES_L, LOREM_IPSUM } from "../../modules/enums.mjs";
import { BKGD_STARLABS, LOGO_STARLABS, STARLABS_LOGO_SM } from "../../modules/media.mjs";

// ==================================================================================================================

export function doAppStart() {
	// REALITY-CHECK: Ensure user has a modern browser
	if (!window.Blob) {
		alert("Unsupported Browser\n\nSorry, but you'll need a modern browser - (Chrome, Firefox, Edge, Opera) - to enable this feature.");
		return;
	} else if (typeof PptxGenJS === "undefined") {
		alert("Oops!\n\n`PptxGenJS` is undefined - maybe a bad link to the 'pptxgen.js' file or something...?\n");
		return;
	}

	// STEP 1: Set UI (if you're me :-)
	if (window.location.href.indexOf("http://localhost:8000/") > -1) {
		$("#tab1sec1").click();
		$("#tab1sec2").click();
	}

	// STEP 2: Show library info
	{
		if (typeof Promise !== "function") {
			$("header").after('<div class="alert alert-danger">Promise is undefined! (IE11 requires promise.min.js)</div>');
		} else {
			let pptx = new PptxGenJS();
			$("#infoBar").append(
				'<div class="col px-0 text-primary"><div class="iconSvg size24 info"></div>Version: <span>' + pptx.version + "</span></div>"
			);
			$("#infoBar").append(
				"<div class='col-auto text-success text-nowrap'><span style='cursor:help' title='" +
					Object.keys(pptx.ChartType).join(" | ") +
					"'><div class='iconSvg size24 circle check'></div>pptx.ChartType = " +
					Object.keys(pptx.ChartType).length +
					"</span></div>"
			);
			$("#infoBar").append(
				"<div class='col-auto text-success text-nowrap'><span style='cursor:help' title='" +
					Object.keys(pptx.SchemeColor).join(" | ") +
					"'><div class='iconSvg size24 circle check'></div>pptx.SchemeColor = " +
					Object.keys(pptx.SchemeColor).length +
					"</span></div>"
			);
			$("#infoBar").append(
				'<div class="col-auto text-success text-nowrap"><span><div class="iconSvg size24 circle check"></div>pptx.ShapeType = ' +
					Object.keys(pptx.ShapeType).length +
					"</span></div>"
			);
		}
	}

	// STEP 3: Build UI elements
	buildDataTable();
	let pptx = new PptxGenJS();
	["MASTER_SLIDE", "THANKS_SLIDE", "TITLE_SLIDE"].forEach(function (name, idx) {
		$("#selSlideMaster").append('<option value="' + name + '">' + name + "</option>");
	});

	// STEP 4: Populate code areas
	{
		$("#demo-basic").text(
			"// STEP 1: Create a new Presentation\n" +
				"let pptx = new PptxGenJS();\n" +
				"\n" +
				"// STEP 2: Add a new Slide to the Presentation\n" +
				"let slide = pptx.addSlide();\n" +
				"\n" +
				"// STEP 3: Add any objects to the Slide (charts, tables, shapes, images, etc.)\n" +
				"slide.addText(\n" +
				"  'BONJOUR - CIAO - GUTEN TAG - HELLO - HOLA - NAMASTE - OLÀ - ZDRAS-TVUY-TE - こんにちは - 你好',\n" +
				"  { x:0.0, y:0.25, w:'100%', h:1.5, align:'center', fontSize:24, color:'0088CC', fill:{ color:'F1F1F1' } }\n" +
				");\n" +
				"\n" +
				"// STEP 4: Send the PPTX Presentation to the user, using your choice of file name\n" +
				"pptx.writeFile({ fileName: 'PptxGenJs-Basic-Slide-Demo' });\n"
		);

		$("#demo-sandbox").html(
			"let pptx = new PptxGenJS();\n" +
				"let slide = pptx.addSlide();\n" +
				//+ "pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });\n"
				//+ "pptx.layout = 'A3';\n"
				"\n" +
				"slide.addText(\n" +
				"  [\n" +
				"    { text:'Did You Know?', options:{ fontSize:48, color:pptx.SchemeColor.accent1, breakLine:true } },\n" +
				"    { text:'writeFile() returns a Promise', options:{ fontSize:24, color:pptx.SchemeColor.accent6, breakLine:true } },\n" +
				"    { text:'!', options:{ fontSize:24, color:pptx.SchemeColor.accent6, breakLine:true } },\n" +
				"    { text:'(pretty cool huh?)', options:{ fontSize:24, color:pptx.SchemeColor.accent3 } }\n" +
				"  ],\n" +
				"  { x:1, y:1, w:'80%', h:3, align:'center', fill:{ color:pptx.SchemeColor.background2, transparency:50 } }\n" +
				");\n" +
				"\n" +
				"pptx.writeFile({ fileName: 'PptxGenJS-Sandbox.pptx' });\n"
		);

		$("#demo-master").html(
			"pptx.defineSlideMaster({\n" +
				"  title : 'MASTER_SLIDE',\n" +
				"  margin: [ 0.5, 0.25, 1.00, 0.25 ],\n" +
				"  background: { color: 'FFFFFF' },\n" +
				"  objects: [\n" +
				"    { image: { x:11.45, y:5.95, w:1.67, h:0.75, data:STARLABS_LOGO_SM } },\n" +
				"    { rect:  { x:0, y:6.9, w:'100%', h:0.6, fill: { color:'003b75' } } },\n" +
				"    { text:  {\n" +
				"        text: 'S.T.A.R. Laboratories - Confidential',\n" +
				"        options: { x:0, y:6.9, w:'100%', align:'center', color:'FFFFFF', fontSize:12 }\n" +
				"    }}\n" +
				//+ "    }},\n"
				//+ "    {placeholder: { options:{ name:'title', type:'title', x:0.5, y:0.2, w:12, h:1.0 }, text:'' }}\n"
				//+ "    {placeholder: { options:{ name:'body', type:'body', x:6.0, y:1.5, w:12, h:5.25 }, text:'' }}\n"
				"  ],\n" +
				"  slideNumber: { x:1.0, y:7.0, color:'FFFFFF' }\n" +
				"});\n"
		);
	}

	// STEP 5: Demo setup
	$("#tabLargeCellText tbody td").text(LOREM_IPSUM.substring(0, 3000));
	for (let idx = 0; idx < 30; idx++) {
		$("#tabLotsOfLines tbody").append("<tr><td>Row-" + idx + "</td><td>Col-B</td><td>Col-C</td></tr>");
	}

	// LAST: Re-highlight code
	$(".tab-content code.language-javascript").each(function (idx, ele) {
		Prism.highlightElement($(ele)[0]);
	});

	// LAST: Nav across sessions
	doNavRestore();
}

export function runAllDemos() {
	if (console.time) console.time("runAllDemos");
	$("#modalBusy").modal("show");

	runEveryTest()
		.catch(function (err) {
			console.error(err.toString());
			$("#modalBusy").modal("hide");
		})
		.then(function () {
			if (console.timeEnd) console.timeEnd("runAllDemos");
			$("#modalBusy").modal("hide");
		});
}

export function execGenSlidesFunc(type) {
	if (console.time) console.time("execGenSlidesFunc: " + type);
	$("#modalBusy").modal("show");

	execGenSlidesFuncs(type)
		.catch(function (err) {
			$("#modalBusy").modal("hide");
			console.error(err);
		})
		.then(function () {
			$("#modalBusy").modal("hide");
			if (console.timeEnd) console.timeEnd("execGenSlidesFunc: " + type);
		});
}

export function buildDataTable() {
	// STEP 1:
	$("#tabAutoPaging tbody").empty();

	// STEP 2:
	for (let idx = 0; idx < $("#numTab2SlideRows").val(); idx++) {
		let strHtml =
			"<tr>" +
			'<td style="text-align:center">' +
			(idx + 1) +
			"</td>" +
			"<td>" +
			TABLE_NAMES_L[Math.floor(Math.random() * 10)] +
			"</td>" +
			"<td>" +
			TABLE_NAMES_F[Math.floor(Math.random() * 10)] +
			"</td>" +
			"<td>Text:<br>" +
			LOREM_IPSUM.substring(0, (Math.floor(Math.random() * 10) + 2) * 130) +
			"</td>" +
			"</tr>";
		$("#tabAutoPaging tbody").append(strHtml);
	}

	// STEP 3: Add some style to table for testing
	// TEST Padding
	$("#tabAutoPaging thead th").css("padding", "10px 5px");
	// TEST font-size/auto-paging
	$("#tabAutoPaging tbody tr:first-child td:last-child").css("font-size", "12px");
	$("#tabAutoPaging tbody tr:last-child td:last-child").css("font-size", "16px");
}

export function table2slidesDemoForTab(inTabId, inOpts) {
	let pptx = new PptxGenJS();
	pptx.tableToSlides(inTabId, inOpts || null);
	pptx.writeFile({ fileName: `${inTabId}_${getTimestamp()}` });
}

export function table2slides1() {
	// FIRST: Instantiate new PptxGenJS instance
	let pptx = new PptxGenJS();

	// STEP 1: Add Master Slide defs / Set slide size/layout
	addMasterDefs(pptx);
	pptx.layout = "LAYOUT_WIDE";

	// STEP 2: Set generated Slide options
	let objOpts = {};
	//objOpts.verbose = true;
	if ($("input[name=radioHead]:checked").val() == "Y") objOpts.autoPageRepeatHeader = true;
	if ($("#checkStartY").prop("checked")) objOpts.autoPageSlideStartY = Number($("#numTab2SlideStartY").val());
	if ($("#selSlideMaster").val()) objOpts.masterSlideName = $("#selSlideMaster").val();
	//console.log(JSON.stringify(objOpts));

	// STEP 3: Pass table to tableToSlides function to produce 1-N slides
	pptx.tableToSlides("tabAutoPaging", objOpts);

	// LAST: Export Presentation
	pptx.writeFile({ fileName: `Table2Slides_MasterSlide_${getTimestamp()}` });
}

export function table2slides2() {
	// FIRST: Instantiate new PptxGenJS instance
	let pptx = new PptxGenJS();

	// STEP 1: Add Master Slide defs / Set slide size/layout
	addMasterDefs(pptx);
	pptx.layout = "LAYOUT_WIDE";

	// STEP 2: Set generated Slide options
	let objOpts = {};
	//objOpts.verbose = true;
	if ($("input[name=radioHead]:checked").val() == "Y") objOpts.addHeaderToEach = true; // TEST: DEPRECATED: addHeaderToEach
	if ($("#checkStartY").prop("checked")) objOpts.newSlideStartY = Number($("#numTab2SlideStartY").val()); // TEST: DEPRECATED: `newSlideStartY`
	if ($("#selSlideMaster").val()) objOpts.masterSlideName = $("#selSlideMaster").val();

	// STEP 3: Add a custom shape (text in this case) to each Slide
	// EXAMPLE: Add any dynamic content to each generated Slide
	// DESC: Add something you cant predefine in a master - like a username/timestamp for each slide, etc.
	// NOTE: You can do this for all other types as well: .addShape(), .addTable() and .addImage()
	objOpts.addText = {
		text: "(dynamic content - ex:user/datestamp)",
		options: { x: 0.1, y: 0.07, color: "0088CC", fontFace: "Arial", fontSize: 12 },
	};

	// STEP 4: Pass table to tableToSlides function to produce 1-N slides
	pptx.tableToSlides("tabAutoPaging", objOpts);

	// LAST: Export Presentation
	pptx.writeFile({ fileName: "Table2Slides_DynamicText" });
}

// ==================================================================================================================

function doNavRestore() {
	$('.nav a[href="' + window.location.href.substring(window.location.href.toLowerCase().indexOf(".html#") + 5) + '"]').tab("show");
}

function getTimestamp() {
	let dateNow = new Date();
	let dateMM = dateNow.getMonth() + 1;
	let dateDD = dateNow.getDate();
	let h = dateNow.getHours();
	let m = dateNow.getMinutes();
	return (
		dateNow.getFullYear() +
		"" +
		(dateMM <= 9 ? "0" + dateMM : dateMM) +
		"" +
		(dateDD <= 9 ? "0" + dateDD : dateDD) +
		(h <= 9 ? "0" + h : h) +
		(m <= 9 ? "0" + m : m)
	);
}

function addMasterDefs(pptx) {
	// 1:
	pptx.defineSlideMaster({
		title: "TITLE_SLIDE",
		background: { data: BKGD_STARLABS },
		objects: [
			{ line: { x: 3.5, y: 1.0, w: 6.0, line: { color: "0088CC", width: 5 } } },
			{ rect: { x: 0.0, y: 5.3, w: "100%", h: 0.75, fill: { color: "F1F1F1" } } },
			{
				text: {
					text: "Global IT & Services :: Status Report",
					options: { x: 3.0, y: 5.3, w: 5.5, h: 0.75, fontFace: "Arial", fontSize: 20, color: "363636", valign: "middle", margin: 0 },
				},
			},
			{ image: { x: 11.3, y: 6.4, w: 1.67, h: 0.75, data: STARLABS_LOGO_SM } },
		],
	});

	// 2:
	pptx.defineSlideMaster({
		title: "MASTER_SLIDE",
		background: { fill: "F1F1F1" },
		slideNumber: { x: 1.0, y: 7.0, color: "FFFFFF" },
		margin: [0.5, 0.25, 1.25, 0.25],
		objects: [
			{ rect: { x: 0.0, y: 6.9, w: "100%", h: 0.6, fill: { color: "003b75" } } },
			{
				text: {
					text: "S.T.A.R. Laboratories",
					options: { x: 0, y: 6.9, w: "100%", h: 0.6, align: "center", valign: "middle", color: "FFFFFF", fontSize: 12 },
				},
			},
		],
	});

	// 3:
	pptx.defineSlideMaster({
		title: "THANKS_SLIDE",
		background: { fill: "36ABFF" },
		objects: [
			{ rect: { x: 0.0, y: 3.4, w: "100%", h: 2.0, fill: { color: "ffffff" } } },
			{
				text: {
					text: "Thank You!",
					options: { x: 0.0, y: 0.9, w: "100%", h: 1, fontFace: "Arial", color: "FFFFFF", fontSize: 60, align: "center" },
				},
			},
			{ image: { x: 4.6, y: 3.5, w: 4, h: 1.8, data: LOGO_STARLABS } },
		],
	});
}

// ==================================================================================================================
// Old, undocumented, legacy tests below
// ==================================================================================================================

function doTestSimple() {
	let pptx = new PptxGenJS();
	let slide = pptx.addSlide();
	let optsTitle = { color: "9F9F9F", marginPt: 3, border: [0, 0, { pt: "1", color: "CFCFCF" }, 0] };

	pptx.layout({ name: "A3", width: 16.5, height: 11.7 });
	slide.slideNumber({ x: 0.5, y: "90%" });
	slide.addTable([[{ text: "Simple Example", options: optsTitle }]], { x: 0.5, y: 0.13, w: 12.5 });

	//slide.addText('Hello World!', { x:0.5, y:0.7, w:6, h:1, color:'0000FF' });
	slide.addText("Hello 45! ", {
		x: 0.5,
		y: 0.5,
		w: 6,
		h: 1,
		fontSize: 36,
		color: "0000FF",
		shadow: { type: "outer", color: "00AAFF", blur: 2, offset: 10, angle: 45, opacity: 0.25 },
	});
	slide.addText("Hello 180!", {
		x: 0.5,
		y: 1.0,
		w: 6,
		h: 1,
		fontSize: 36,
		color: "0000FF",
		shadow: { type: "outer", color: "ceAA00", blur: 2, offset: 10, angle: 180, opacity: 0.5 },
	});
	slide.addText("Hello 355!", {
		x: 0.5,
		y: 1.5,
		w: 6,
		h: 1,
		fontSize: 36,
		color: "0000FF",
		shadow: { type: "outer", color: "aaAA33", blur: 2, offset: 10, angle: 355, opacity: 0.75 },
	});

	// Bullet Test: Number
	slide.addText(999, { x: 0.5, y: 2.0, w: "50%", h: 1, color: "0000DE", bullet: true });
	// Bullet Test: Text test
	slide.addText("Bullet text", { x: 0.5, y: 2.5, w: "50%", h: 1, color: "00AA00", bullet: true });
	// Bullet Test: Multi-line text test
	slide.addText("Line 1\nLine 2\nLine 3", { x: 0.5, y: 3.5, w: "50%", h: 1, color: "AACD00", bullet: true });

	// Table cell margin:0
	slide.addTable([["margin:0"]], { x: 0.5, y: 1.1, margin: 0, w: 0.75, fill: { color: "FFFCCC" } });

	// Fine-grained Formatting/word-level/line-level Formatting
	slide.addText(
		[
			{ text: "right line", options: { fontSize: 24, fontFace: "Courier New", color: "99ABCC", align: "right", breakLine: true } },
			{ text: "ctr line", options: { fontSize: 36, fontFace: "Arial", color: "FFFF00", align: "center", breakLine: true } },
			{ text: "left line", options: { fontSize: 48, fontFace: "Verdana", color: "0088CC", align: "left" } },
		],
		{ x: 0.5, y: 3.0, w: 8.5, h: 4, margin: 0.1, fill: { color: "232323" } }
	);

	// Export:
	pptx.writeFile({ fileName: "Sample Presentation" });
}

/* The "Text" demo on the PptxGenJS homepage - codified here so we can quickly reproduce the screencaps, etc. as needed */
function doHomepageDemo_Text() {
	let pptx = new PptxGenJS();
	pptx.layout = "LAYOUT_WIDE";
	let slide = pptx.addSlide();

	slide.addText("BONJOUR - CIAO - GUTEN TAG - HELLO - HOLA - \nNAMASTE - OLÀ - ZDRAS-TVUY-TE - こんにちは - 你好", {
		x: 0.0,
		y: 0.0,
		w: "100%",
		h: 1.25,
		align: "center",
		fontSize: 18,
		color: "0088CC",
		fill: { color: "F1F1F1" },
	});

	slide.addText("Line-Level Formatting:", { x: 0.5, y: 1.5, w: "40%", h: 0.38, color: "0088CC" });
	slide.addText(
		[
			{ text: "1st line", options: { fontSize: 24, fontFace: "Courier New", color: "99ABCC", align: "right", breakLine: true } },
			{ text: "2nd line", options: { fontSize: 36, fontFace: "Arial", color: "FFFF00", align: "center", breakLine: true } },
			{ text: "3rd line", options: { fontSize: 48, fontFace: "Verdana", color: "0088CC", align: "left" } },
		],
		{ x: 0.5, y: 2.0, w: 6, h: 2.25, margin: 0.1, fill: { color: "232323" } }
	);

	slide.addText("Bullets: Normal", { x: 8.0, y: 1.5, w: "40%", h: 0.38, color: "0088CC" });
	slide.addText("Line 1\nLine 2\nLine 3", {
		x: 8.0,
		y: 2.0,
		w: "30%",
		h: 1,
		color: "393939",
		fontSize: 16,
		fill: { color: "F2F2F2" },
		bullet: true,
	});

	slide.addText("Bullets: Numbered", { x: 8.0, y: 3.4, w: "40%", h: 0.38, color: "0088CC" });
	slide.addText("Line 1\nLine 2\nLine 3", {
		x: 8.0,
		y: 3.9,
		w: "30%",
		h: 1,
		color: "393939",
		fontSize: 16,
		fill: { color: "F2F2F2" },
		bullet: { type: "number" },
	});

	slide.addText("Bullets: Custom", { x: 8.0, y: 5.3, w: "40%", h: 0.38, color: "0088CC" });
	slide.addText("Star bullet! ", { x: 8.0, y: 5.6, w: "40%", h: 0.38, color: "CC0000", bullet: { code: "2605" } });
	slide.addText("Check bullet!", { x: 8.0, y: 5.9, w: "40%", h: 0.38, color: "00CD00", bullet: { code: "2713" } });

	let shadowOpts = { type: "outer", color: "696969", blur: 3, offset: 10, angle: 45, opacity: 0.8 };
	slide.addText("Text Shadow:", { x: 0.5, y: 6.0, w: "40%", h: 0.38, color: "0088CC" });
	slide.addText("Outer Shadow (blur:3, offset:10, angle:45, opacity:80%)", {
		x: 0.5,
		y: 6.4,
		w: 12,
		h: 0.6,
		fontSize: 32,
		color: "0088cc",
		shadow: shadowOpts,
	});

	pptx.writeFile({ fileName: "Demo-Text" });
}

function testTTS() {
	let pptx = new PptxGenJS();
	pptx.layout = "LAYOUT_WIDE";
	/*
	let slide = pptx.addSlide();
	slide.addText('Table Paging Logic Check', { x:0.0, y:'90%', w:'100%', align:'center', fontSize:18, color:'0088CC', fill:{color:'F2F9FC'} });
	let numMargin = 1.25;
	slide.addShape(pptx.shapes.RECTANGLE, { x:0.0, y:0.0, w:numMargin, h:numMargin, fill:{color:'FFFCCC'} });
	slide.addTable( ['short','table','whatever'], {x:numMargin, y:numMargin, margin:numMargin, colW:2.5, fill:{color:'F1F1F1'}} );
	*/

	// Mimic "table2slides1()"
	// addMasterDefs(pptx);

	// TEST
	//pptx.tableToSlides('tabAutoPaging');
	pptx.tableToSlides("tabAutoPaging", { verbose: true, autoPageRepeatHeader: true /*, autoPageSlideStartY:2*/ });

	pptx.writeFile({ fileName: `PptxGenJs_TTSTest_${getTimestamp()}` });
}

function testTTSMulti() {
	let ttsTitleText = { fontSize: 14, color: "0088CC", bold: true };
	let ttsMultiOpts = { fontSize: 13, color: "9F9F9F", verbose: true };
	let arrRows = [];
	let arrText = [];
	//
	let pptx = new PptxGenJS();
	pptx.layout = "LAYOUT_WIDE";

	for (let idx = 0; idx < TABLE_NAMES_F.length; idx++) {
		let strText = idx == 0 ? LOREM_IPSUM.substring(0, 100) : LOREM_IPSUM.substring(idx * 100, idx * 200);
		arrRows.push([idx, TABLE_NAMES_F[idx], strText]);
		arrText.push([strText]);
	}

	// autoPageLineWeight option demos
	let slide = pptx.addSlide();
	slide.addText(
		[
			{ text: "Table Examples: ", options: ttsTitleText },
			{ text: "autoPageLineWeight:0", options: ttsMultiOpts },
		],
		{ x: 0.5, y: 0.13, w: 3 }
	);
	slide.addTable(arrText, { x: 0.5, y: 0.6, w: 4, margin: 5, border: "CFCFCF", autoPage: true });

	slide.addText(
		[
			{ text: "Table Examples: ", options: ttsTitleText },
			{ text: "autoPageLineWeight:0.5", options: ttsMultiOpts },
		],
		{ x: 5.0, y: 0.13, w: 3 }
	);
	slide.addTable(arrText, { x: 4.75, y: 0.6, w: 4, margin: 5, border: "CFCFCF", autoPage: true, autoPageLineWeight: 0.5 });

	slide.addText(
		[
			{ text: "Table Examples: ", options: ttsTitleText },
			{ text: "autoPageLineWeight:-0.5", options: ttsMultiOpts },
		],
		{ x: 9.0, y: 0.13, w: 3 }
	);
	slide.addTable(arrText, { x: 9.1, y: 0.6, w: 4, margin: 5, border: "CFCFCF", autoPage: true, autoPageLineWeight: -0.5 });

	pptx.writeFile({ fileName: `PptxGenJS_TTSMulti_${getTimestamp()}` });
}

function table2slidesBullets() {
	let pptx = new PptxGenJS();
	pptx.tableToSlides("tableWithBullets");
	pptx.writeFile({ fileName: `tabBullets_${getTimestamp()}` });
}

/* DESC: Test for backward compatibility with Slide Masters defined in `pptxgen.masters.js` */
function testOnly_LegacyMasterSlides() {
	// TEST-ONLY: DO NOT USE/COPY ME!!
	let pptx = new PptxGenJS();
	pptx.layout = "LAYOUT_WIDE";
	let slide = pptx.addSlide(pptx.masters.TITLE_SLIDE);
	pptx.writeFile({ fileName: `Demo-LegacyMasterSlides_${getTimestamp()}` });
}
