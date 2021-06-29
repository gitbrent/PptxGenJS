import React from "react";
import pptxgen from "pptxgenjs"; // react-app webpack will use package.json `"module": "dist/pptxgen.es.js"` value
import { testMainMethods, testTableMethod } from "./tstest/Test";
import logo from "./logo.svg";
import "./App.css";

const demoCode = `import pptxgen from "pptxgenjs";

let pptx = new pptxgen();
let slide = pptx.addSlide();

slide.addText(
  "React Demo!",
  { x:1, y:0.5, w:'80%', h:1, fontSize:36, align:'center', fill:{ color:'D3E3F3' }, color:'008899' }
);

let dataChartSunburst = [
	{
		name: "Datenreihe 1",
		labels: ["Arch Stack", "Arch Stack", "45", "32",
			"Arch Stack", "Arch Stack", "45", "13",
			"Arch Stack", "66", "66", "66",
			"BC 2", "BC 2", "", "54",
			"BC 2", "BC 2", "3", "3",
			"Weitere BC", "13", "", ""],
		values: [32, 13, 66, 54, 3, 13]
	},
];

slide.addChart(pptx.ChartType.sunburst, dataChartSunburst, { x: 1, y: 1.9, w: 8, h: 3,
	chartColors:  ["", "", "ff0000", "11e2ff", "ff0000", "b5fd20", "a7d6d4", "b5fd20", "ffbd11", "", "", "20ff47", "", "", "ff0000", "", ""]
});

slide.addText(
  "PpptxGenJS version:",
  { x:0, y:5.3, w:'100%', h:0.33, align:'center', fill:{ color:'E1E1E1' }, color:'A1A1A1' }
);

pptx.writeFile({ fileName: 'pptxgenjs-demo-react.pptx' });`;

function App() {
	function runDemo() {
		let pptx = new pptxgen();
		let slide = pptx.addSlide();

		let dataChartPie = [
			{
				name: "Region 1",
				labels: ["May", "June", "July", "August", "September"],
				values: [26, 53, 100, 75, 41],
			},
		];

		//slide.addShape(pptx.ShapeType.rect, { x: 4.36, y: 2.36, w: 5, h: 2.5, fill: pptx.SchemeColor.background2 });

		//slide.addText("React Demo!", { x: 1, y: 1, w: "80%", h: 1, fontSize: 36, fill: "eeeeee", align: "center" });
		/*slide.addText("React Demo!", {
			x: 1,
			y: 0.5,
			w: "80%",
			h: 1,
			fontSize: 36,
			align: 'center',
			fill: { color:'D3E3F3' },
			color: '008899',
		});*/

		// example from Windows PPT
		let dataChartSunburstWindowsExample = [
			{
				name: "Umsatz",
				labels: ["1.", "Jan", "",
					"", "Feb", "Woche 1",
					"", "", "Woche 2",
					"", "", "Woche 3",
					"", "", "Woche 4",
					"", "Mrz", "",
					"2.", "Apr", "",
					"", "Mai", "",
					"", "Juni", "",
					"3.", "Jul", "",
					"", "Aug", "",
					"", "Sep", "",
					"4.", "Okt", "",
					"", "Nov", "",
					"", "Dez", ""],
				values: [3.5, 1.2, 0.8, 0.6, 0.5, 1.7, 1.1, 0.8, 0.3, 0.7, 0.6, 0.1, 0.5, 0.4, 0.3],
				sizes: [3]
			}, {
				name: "textColors",
				labels: ["ffffff", "ffffff", "",
					"", "ffffff", "ffffff",
					"", "", "ffffff",
					"", "", "ffffff",
					"", "", "ffffff",
					"", "ffffff", "",
					"ffffff", "ffffff", "",
					"", "ffffff", "",
					"", "ffffff", "",
					"131313", "131313", "",
					"", "131313", "",
					"", "131313", "",
					"ffffff", "ffffff", "",
					"", "ffffff", "",
					"", "ffffff", ""], // for every slice 1 color in order of labels with empty entries
				sizes: [3]
			}, {
				name: "borderColors",
				labels: ["cccccc", "cccccc",
					"cccccc", "cccccc",
					"cccccc",
					"cccccc",
					"cccccc",
					"cccccc",
					"cccccc", "cccccc",
					"cccccc",
					"cccccc",
					"131313", "131313",
					"131313",
					"131313",
					"cccccc", "cccccc",
					"cccccc",
					"cccccc"], // for every slice 1 color in order of labels without empty cells
			}];
		let chartColorsWindowsExample = ["354567", "354567",
			"354567", "354567",
			"354567",
			"354567",
			"354567",
			"354567",
			"3a87ad", "3a87ad",
			"3a87ad",
			"3a87ad",
			"ffffff", "ffffff",
			"ffffff",
			"ffffff",
			"ebad60", "ebad60",
			"ebad60",
			"ebad60"]

		// example for sunburst "Landscape"
		let dataLandscape = [{
			name: "Items",
			labels: ["49", "49", "", "", "", "",
				"4", "4", "", "", "", "",
				"2", "2", "2", "2", "2", "2",
				"", "", ".2", ".2", ".2", ".2",
				"", "", "..2", "..2", "..2", "..2",
				"", "", "", "", "...2", "...2",
				"13", "13", "", "", "", "",
				"5", "5", "3", "3", "4", "4",
				"", "", "", "", "6", "6",
				"", "", "", "", "1", "1",
				"", "", "", "", "3", "3",
				"", "", "", "", "5", "5",
				"", "", "", "", "29", "29",
				"", "", "", "", "1", "1",
				"38", "38", "", "", "", "",
				"6", "6", "", "", "", "",
				"1", "1", "", "", "", "",
				"20", "20", "", "", "", "",
				".1", ".1", "", "", "", "",
				".13", "13", "7", "7", "", "",
				"", "", ".3", ".3", "", "",
				"", "", "..3", "..3", "", "",
				"10", "10", "", "", "", "",
				".5", ".5", "1", "1", "", "",
				"", "", "9", "9", "", "",
				"", "", "8", "8", "12", "12",
				"", "", "", "", "7", "7",
				"", "", "", "", "8", "8",
				"", "", "", "", "10", "10",
				"", "", "", "", "22", "22",
				"", "", "", "", "18", "18",
				"", "", "", "", "9", "9",
				"", "", "6", "6", "", "",
				"43", "43", "", "", "", "",
				".13", ".13", "", "", "", "",
				"12", "12", "", "", "", "",
				"8", "8", "", "", "", "",
				".6", ".6", "", "", "", "",
				".2", ".2", "", "", "", "",
				"7", "7", "", "", "", "",
				"227", "227", "", "", "", ""],
			values: ["49", "4", "2", "2", "2", "2", "13", "4", "6", "1", "3", "5", "29", "1", "38", "6", "1", "20", "1", "7", "3", "3", "10", "1", "9", "14", "3", "23", "5", "8", "13", "34", "6", "43", "13", "12", "8", "6", "2", "7", "227"],
			sizes: [6]
		}];
		let colorsLandscape = ["354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad", "354567", "3a87ad", "354567", "3a87ad",
			"354567", "3a87ad", "354567", "3a87ad",
			"354567", "3a87ad", "354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad", "354567", "3a87ad", "354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad", "354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad", "354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad", "354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad",
			"354567", "3a87ad"];
		let optsLandscape = {
			x: 1.36,
			y: 1.0,
			h: '80%',
			showTitle: false,
			showLegend: false,
			dataLabelFontSize: 10,
			dataLabelFontFace: 'Calibri',
			chartColors: colorsLandscape
		};

		// example for sunburst "Hierarchie"
		let dataHierarchy = [{
			name: 'Hierarchy',
			 labels: ["Root 1","","",
				 "Root 1","Root 1 / Eine sehr sehr sehr laaaaaaaaaaaaaaaaaaaaaaaaaaange Namen-Business Capability die über mehrere Zeilen dargestellt werden muss ü &","",
				 "Root 1","Root 1 / Eine sehr sehr sehr laaaaaaaaaaaaaaaaaaaaaaaaaaange Namen-Business Capability die über mehrere Zeilen dargestellt werden muss ü &","Leäf & 1.1.1",
				 "Root 1","Root 1 / Eine sehr sehr sehr laaaaaaaaaaaaaaaaaaaaaaaaaaange Namen-Business Capability die über mehrere Zeilen dargestellt werden muss ü &","Eine sehr sehr sehr laaaaaaaaaaaaaaaaaaaaaaaaaaange Namen-Business Capability die über mehrere Zeilen dargestellt werden muss / Eine sehr sehr sehr laaaaaaaaaaaaaaaaaaaaaaaaaaange Namen-Business Capability als Kind die über mehrere Zeilen dargestellt werden muss",
				 "Root 1","Root 1 / Eine sehr sehr sehr laaaaaaaaaaaaaaaaaaaaaaaaaaange Namen-Business Capability die über mehrere Zeilen dargestellt werden muss ü &","Leaf 1.1.3",
				 "Root 1","Root 1 / Eine sehr sehr sehr laaaaaaaaaaaaaaaaaaaaaaaaaaange Namen-Business Capability die über mehrere Zeilen dargestellt werden muss ü &","Leaf 1.1.4",
				 "Root 1","Root 1 / Eine sehr sehr sehr laaaaaaaaaaaaaaaaaaaaaaaaaaange Namen-Business Capability die über mehrere Zeilen dargestellt werden muss ü &","Leaf 1.1.5",
				 "Root 1","Root 1 / Eine sehr sehr sehr laaaaaaaaaaaaaaaaaaaaaaaaaaange Namen-Business Capability die über mehrere Zeilen dargestellt werden muss ü &","Leaf 1.1.6",
				 "Sales & Marketing","","",
				 "Sales & Marketing","Sales & Marketing / Node 2.1","",
				 "Sales & Marketing","Sales & Marketing / Node 2.1","Leaf 2.1.1",
				 "Sales & Marketing","Sales & Marketing / Node 2.1","Leaf 2.1.2",
				 "Sales & Marketing","Sales & Marketing / Node 2.2","",
				 "Sales & Marketing","Sales & Marketing / Node 2.2","Leaf 2.2.1",
				 "Sales & Marketing","Sales & Marketing / Node 2.2","Leaf 2.2.2",
				 "Root 3","","",
				 "Root 3","Leaf 3.1","",
				 "Root 4","","",
				 "Root 4","Node 4.1","",
				 "Root 4","Node 4.1","Leaf 4.1.1",
				 "Root 4","Node 4.1","Leaf 4.1.2",
				 "Root 4","Node 4.2","",
				 "Root 4","Node 4.2","Leaf 4.2.1",
				 "Root 4","Leaf 4.1","",
				 "Root 4","Node 4.3","",
				 "Root 4","Node 4.3","Leaf 4.3.1",
				 "Root 4","Node 4.3","Leaf 4.3.2",
				 "Root 5","","",
				 "Root 5","Leaf 5.1","",
				 "Root 6","","",
				 "Root 6","Node 6.1","",
				 "Root 6","Node 6.1","Leaf 6.1.1",
				 "Root 6","Node 6.1","Leaf 6.1.2",
				 "Root 6","Node 6.1","Leaf 6.1.3",
				 "Root 7","","",
				 "Root 7","Leaf 7.1","",
				 "Root 7","Leaf 7.2","",
				 "Root 7","Leaf 7.3","",
				 "Root 7","Leaf 7.4","",
				 "Root 7","Leaf 7.5","",
				 "Root 7","Node 7.1","",
				 "Root 7","Node 7.1","Leaf 7.1.1",
				 "Root 7","Node 7.1","Leaf 7.1.2",
				 "Root 7","Node 7.2","",
				 "Root 7","Node 7.2","Leaf 7.2.1",
				 "Root 7","Leaf 7.6","",
				 "Root 7","Node 7.3","",
				 "Root 7","Node 7.3","Leaf 7.3.1",
				 "Root 7","Node 7.3","Leaf 7.3.2",
				 "Root 7","Leaf 7.7","",
				 "Root 7","Node 7.4","",
				 "Root 7","Node 7.4","Leaf 7.4.1"],
			values: ["-10", "-6", "1", "1", "1", "1", "1", "1", "-4", "-2", "1", "1", "-2", "1", "1", "-1", "1", "-6", "-2", "1",
					 "1", "-1", "1", "1", "-2", "1", "1", "-1", "1", "-3", "-3", "1", "1", "1", "-13", "1", "1", "1", "1", "1", "-2",
					 "1", "1", "-1", "1", "1", "-2", "1", "1", "1", "-1", "1"],
			sizes: [3]
		}]
		let colorsHierarchy = ["ffffff",
			"496a8f",
			"ffffff",
			"ffffff",
			"496a8f",
			"ffffff",
			"ffffff",
			"ffffff",
			"ff0000",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff",
			"ffffff"]
		slide.addChart(pptx.charts.SUNBURST, dataHierarchy, {
			x: 0.8,
			y: 1.0,
			w: 4.0,
			h: 4.0,
			showTitle: false,
			showLegend: false,
			showLabel: true,
			showValue: false,
			dataLabelFontSize: 5,
			dataLabelFontFace: 'Calibri',
			chartColors: colorsHierarchy
		});

/*		slide.addChart(pptx.charts.SUNBURST, dataLandscape, {
			x: 6,
			y: 1.25,
			w: 4.0,
			h: 4.0,
			showTitle: false,
			showLegend: false,
			showValue: true,
			dataLabelFontSize: 10,
			dataLabelFontFace: 'Calibri',
			chartColors: colorsLandscape
		});*/

		//slide.addChart(pptx.ChartType.pie, dataChartPie, { x: 1, y: 1.25, w: 4.0, h: 4.0 });

		// example ala Michi
		let dataChartSunburstMichi = [
			{
				name: "Datenreihe 1",
				labels: ['Root 1', '', '',
					'Root 1', 'Node 1.1', '',
					'Root 1', 'Node 1.1', 'Leaf 1.1.1',
					'Root 1', 'Node 1.1', 'Leaf 1.1.2',
					'Root 1', 'Leaf 1.1', '',
					'Root 1', 'Leaf 1.2', '',
					'Root 2', '', '',
					'Root 2', 'Node 2.1', '',
					'Root 2', 'Node 2.1', 'Leaf 2.1.1',
					'Root 2', 'Leaf 2.1', '',
					'Root 2', 'Node 2.2', '',
					'Root 2', 'Node 2.2', 'Leaf 2.2.1',
					'Root 2', 'Node 2.2', 'Leaf 2.2.2',
					'Root 2', 'Node 2.2', 'Leaf 2.2.3',
					'Leaf 1', '', ''],
				values: [-234, -119, 60, 59, 58, 57, -270, -56, 56, 55, -159, 54, 53, 52, 51],
				sizes: [3]
			}
		];
		let chartColorsMichi = ['354567',
			'3a87ad',
			'354567',
			'354567',
			'3a87ad',
			'3a87ad',
			'354567',
			'3a87ad',
			'354567',
			'3a87ad',
			'3a87ad',
			'354567',
			'354567',
			'354567',
			'354567'] // for every slice 1 color in order of labels without empty cells

		slide.addChart(pptx.charts.SUNBURST, dataChartSunburstMichi, {
			x: 5.0,
			y: 1.0,
			w: 4.0,
			h: 4.0,
			showTitle: false,
			showLegend: false,
			showValue: true,
			dataLabelFontSize: 10,
			dataLabelFontFace: 'Calibri',
			chartColors: chartColorsMichi
		});

		/*
                // example with 20 rings
                let dataChartSunburst = [
                    {
                        name: "Datenreihe 1",
                        labels: ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19",
                            "0", "01", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39",
                            "0", "", "", "", "", "45", "", "", "", "", "50", "", "", "", "", "55", "", "", "", ""],
                        values: [33, 12, 55],
                        sizes: [20]
                    }, {
                        name: "uniqueLabels",
                        labels: ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19",
                            "", "01", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39",
                            "", "", "", "", "", "45", "", "", "", "", "50", "", "", "", "", "55", "", "", "", ""],
                        sizes: [20]
                    },
                    {
                        name: "textColors",
                        labels: ["cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc",
                            "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc",
                            "cccccc", "", "", "", "", "cccccc", "", "", "", "", "cccccc", "", "", "", "", "cccccc", "", "", "", ""], // for every slice 1 color in order of labels
                        sizes: [20]
                    },
                    {
                        name: "borderColors",
                        labels: ["cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc",
                            "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc", "cccccc",
                            "cccccc", "cccccc", "cccccc"], // for every slice 1 color in order of labels without empty cells
                    }
                ];
                let chartColors =  ["cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000",
                    "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000", "cccccc", "ff0000",
                    "ff0000", "cccccc", "ff0000"] // for every slice 1 color in order of labels without empty cells
        */

		/*slide.addText(`PpptxGenJS version: ${pptx.version}`, {
			x: 0,
			y: 5.3,
			w: "100%",
			h: 0.33,
			fontSize: 10,
			align: 'center',
			fill: 'E1E1E1', //{ color: pptx.SchemeColor.background2 },
			color: 'A1A1A1' // pptx.SchemeColor.accent3,
		});*/

		pptx.writeFile({ fileName: "pptxgenjs-demo-react.pptx" });
	}

	return (
		<div>
			<nav className="navbar navbar-expand-lg navbar-dark bg-primary">
				<a className="navbar-brand" href="https://gitbrent.github.io/PptxGenJS/">
					<img src={logo} width="30" height="30" className="d-inline-block align-top mr-2" alt="logo" />
					PptxGenJS
				</a>
				<button
					className="navbar-toggler"
					type="button"
					data-toggle="collapse"
					data-target="#navbarColor01"
					aria-controls="navbarColor01"
					aria-expanded="false"
					aria-label="Toggle navigation"
				>
					<span className="navbar-toggler-icon"></span>
				</button>

				<div className="collapse navbar-collapse" id="navbarColor01">
					<ul className="navbar-nav mr-auto">
						<li className="nav-item active">
							<a className="nav-link" href="https://gitbrent.github.io/PptxGenJS/demo-react/index.html">
								Home <span className="sr-only">(current)</span>
							</a>
						</li>
					</ul>
					<form className="form-inline my-2 my-lg-0">
						<button
							type="button"
							className="btn btn-outline-info mx-3 my-2 my-sm-0"
							onClick={(ev) => {
								window.open("https://gitbrent.github.io/PptxGenJS/demo/", true);
							}}
						>
							Demo Page
						</button>
						<button
							type="button"
							className="btn btn-outline-info mx-3 my-2 my-sm-0"
							onClick={(ev) => {
								window.open("https://github.com/gitbrent/PptxGenJS", true);
							}}
						>
							GitHub Project
						</button>
						<button
							type="button"
							className="btn btn-outline-info mx-3 my-2 my-sm-0"
							onClick={(ev) => {
								window.open("https://gitbrent.github.io/PptxGenJS/docs/installation.html", true);
							}}
						>
							API Docs
						</button>
					</form>
				</div>
			</nav>

			<main className="container">
				<div className="jumbotron mt-5">
					<h1 className="display-4">React Demo</h1>
					<p className="lead">Sample React application to demonstrate using the PptxGenJS library as a module.</p>
					<hr className="my-4" />

					<h5 className="text-info">Demo Code (.tsx)</h5>
					<pre className="my-4">
						<code className="language-javascript">{demoCode}</code>
					</pre>

					<div className="row">
						<div className="col">
							<button type="button" className="btn btn-success w-100 mr-3" onClick={(_ev) => runDemo()}>
								Run Demo
							</button>
						</div>
						<div className="col">
							<button type="button" className="btn btn-primary w-100" onClick={(_ev) => testMainMethods()}>
								Run Std Tests
							</button>
						</div>
						<div className="col">
							<button type="button" className="btn btn-primary w-100" onClick={(_ev) => testTableMethod()}>
								Run HTML2PPT Test
							</button>
						</div>
					</div>

					<table id="html2ppt" className="table table-dark" style={{ display: "none" }}>
						<thead>
							<tr>
								<th>col 1</th>
								<th>col 2</th>
								<th>col 3</th>
							</tr>
						</thead>
						<tbody>
							<tr>
								<td>cell 1</td>
								<td>cell 2</td>
								<td>cell 3</td>
							</tr>
						</tbody>
					</table>
				</div>
			</main>
		</div>
	);
}

export default App;
