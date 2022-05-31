/**
 * NAME: data_convert.mjs
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DESC: Convert Excel-style, tab-delimited data to a default pptxgenjs chart
 * VER.: 1.0.0
 * BLD.: 20220530
 */

const INPUT_ELE_ID = "dataConvTextArea";
/* INPUT:
	|----|---------|-------|------|
	|    |Australia|Belgium|Canada|
	|1976|    10.25|  21.54| 18.07|
	|1977|    10.23|  25.39| 19.21|
	|1978|    10.80|  28.35| 20.62|
	|1979|    12.82|  30.45| 23.89|
	|1980|    13.64|  30.52| 25.02|
*/
/* OUTPUT:
	const LABELS = ["1976", "1977", "1978", "1979", "1980"];
	const CHART_DATA = [
		{ name: "Australia", labels: LABELS, values: [5.01, 4.59, 3.65, 3.62, 3.14] },
		{ name: "Belgium",   labels: LABELS, values: [4.63, 3.67, 3.26, 3.21, 2.79] },
		{ name: "Canada",    labels: LABELS, values: [4.27, 3.61, 3.23, 3.24, 2.78] },
	];
*/
function convertInputData() {
	const tabDelimData = document.getElementById(INPUT_ELE_ID).value;

	let dataLabels = [];
	let chartData = [];

	// 1: build data
	tabDelimData
		.split("\n")
		.filter((_item, idx) => idx === 0)
		.forEach((row, _idx) => {
			row.split("\t").forEach((cell, idx) => {
				if (idx > 0) {
					chartData.push({
						name: cell,
						labels: "dataLabels",
						values: [],
					});
				}
			});
		});

	// 2: add data
	tabDelimData
		.split("\n")
		.filter((_item, idx) => idx > 0)
		.forEach((row, idx) => {
			row.split("\t").forEach((cell, idy) => {
				if (idy === 0) {
					dataLabels.push(cell);
				} else {
					chartData[idy - 1].values.push(Number(cell));
				}
			});
		});

	// 3: show results
	document.getElementById(INPUT_ELE_ID).value =
		`const DATA_LABELS = ${JSON.stringify(dataLabels, "", 4)};\n` +
		`const CHART_DATA = ${JSON.stringify(chartData, "", 4).replace(/\"dataLabels\"/gi, "DATA_LABELS")};\n` +
		`slide.addChart(pptx.charts.LINE, CHART_DATA, {x:0.25, y:0.25, w:'95%', h:'90%', showLegend:true});`;
}

export function generateUI(parentEleID) {
	// 1
	const mainUI = document.createElement("div");

	// 2
	const newTitle = document.createElement("h5");
	newTitle.classList.add("text-primary");
	newTitle.textContent = "Chart Data Importer";
	mainUI.appendChild(newTitle);

	// 3
	const newTextarea = document.createElement("TEXTAREA");
	newTextarea.classList.add("form-control", "font-monospace", "text-sm", "w-100");
	newTextarea.id = INPUT_ELE_ID;
	newTextarea.rows = 20;
	mainUI.appendChild(newTextarea);

	// 4
	const newDesc = document.createElement("div");
	newDesc.classList.add("text-sm", "text-muted", "my-3");
	newDesc.textContent = "(paste tab-delimited data here) (aka: copy-n-paste from Excel)";
	mainUI.appendChild(newDesc);

	// 5
	const newBtn = document.createElement("button");
	newBtn.type = "button";
	newBtn.classList.add("btn", "btn-primary", "px-5", "me-3");
	newBtn.textContent = "Convert Data";
	newBtn.onclick = convertInputData;
	mainUI.appendChild(newBtn);

	// 6
	const newBtnCode = document.createElement("button");
	newBtnCode.type = "button";
	newBtnCode.classList.add("btn", "btn-secondary", "ms-3");
	newBtnCode.textContent = "Show Sandbox";
	newBtnCode.onclick = function () {
		document.getElementById("codeSandbox").classList.remove("d-none");
	};
	mainUI.appendChild(newBtnCode);

	// LAST
	document.getElementById(parentEleID).append(mainUI);
}
