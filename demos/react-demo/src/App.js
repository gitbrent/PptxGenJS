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

slide.addChart(
  pptx.ChartType.radar, dataChartRadar, { x:1.0, y:1.9, w:8, h:3 }
);

slide.addText(
  "PpptxGenJS version:",
  { x:0, y:5.3, w:'100%', h:0.33, align:'center', fill:{ color:'E1E1E1' }, color:'A1A1A1' }
);

pptx.writeFile("pptxgenjs-demo-react.pptx");`;

function App() {
	function runDemo() {
		let pptx = new pptxgen();
		let slide = pptx.addSlide();

		let dataChartRadar = [
			{
				name: "Region 1",
				labels: ["May", "June", "July", "August", "September"],
				values: [26, 53, 100, 75, 41],
			},
		];
		//slide.addChart(pptx.ChartType.radar, dataChartRadar, { x: 0.36, y: 2.25, w: 4.0, h: 4.0, radarStyle: "standard" });

		//slide.addShape(pptx.ShapeType.rect, { x: 4.36, y: 2.36, w: 5, h: 2.5, fill: pptx.SchemeColor.background2 });

		//slide.addText("React Demo!", { x: 1, y: 1, w: "80%", h: 1, fontSize: 36, fill: "eeeeee", align: "center" });
		slide.addText("React Demo!", {
			x: 1,
			y: 0.5,
			w: "80%",
			h: 1,
			fontSize: 36,
			align: 'center',
			fill: { color:'D3E3F3' },
			color: '008899',
		});

		slide.addChart(pptx.ChartType.radar, dataChartRadar, { x: 1, y: 1.9, w: 8, h: 3 });

		slide.addText(`PpptxGenJS version: ${pptx.version}`, {
			x: 0,
			y: 5.3,
			w: "100%",
			h: 0.33,
			fontSize: 10,
			align: 'center',
			fill: 'E1E1E1', //{ color: pptx.SchemeColor.background2 },
			color: 'A1A1A1' // pptx.SchemeColor.accent3,
		});

		pptx.writeFile("pptxgenjs-demo-react.pptx");
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
