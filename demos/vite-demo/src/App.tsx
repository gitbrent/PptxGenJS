// NOTE: previous {create-react-app} is webpack-based and will use package.json `module: "dist/pptxgen.es.js"` value
// NOTE: this Vite+React demo is using `main: "dist/pptxgen.cjs.js"` value, so we hard-code below to TEST
/* // @ts-expect-error (manually import the es module for TESTING!) */
//import pptxgen from "pptxgenjs/dist/pptxgen.cjs.js";
import pptxgen from "pptxgenjs";
import { testMainMethods, testTableMethod } from "./tstest/Test";
import { demoCode } from "./enums";
import logo from "./assets/logo.png";
import './scss/styles.scss';

function App() {
	function runDemo() {
		const pptx = new pptxgen();
		const slide = pptx.addSlide();

		const dataChartRadar = [
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
			align: "center",
			fill: { color: "D3E3F3" },
			color: "008899",
		});

		slide.addChart(pptx.ChartType.radar, dataChartRadar, { x: 1, y: 1.9, w: 8, h: 3 });

		slide.addText(`PpptxGenJS version: ${pptx.version}`, {
			x: 0,
			y: 5.3,
			w: "100%",
			h: 0.33,
			fontSize: 10,
			align: "center",
			fill: { color: "E1E1E1" }, //{ color: pptx.SchemeColor.background2 },
			color: "A1A1A1", // pptx.SchemeColor.accent3,
		});

		pptx.writeFile({ fileName: "pptxgenjs-demo-react.pptx" });
	}

	const htmlNav = () => {
		return <nav className="navbar navbar-expand-lg bg-primary" data-bs-theme="dark">
			<div className="container-fluid">
				<a className="navbar-brand" href="https://gitbrent.github.io/PptxGenJS/">
					<img src={logo} alt="logo" width="30" height="30" className="d-inline-block align-text-center me-2" />
					PptxGenJS
				</a>
				<button className="navbar-toggler" type="button"
					data-bs-toggle="collapse"
					data-bs-target="#navbarText"
					aria-controls="navbarText"
					aria-expanded="false"
					aria-label="Toggle navigation"
				>
					<span className="navbar-toggler-icon"></span>
				</button>
				<div className="collapse navbar-collapse" id="navbarText">
					<ul className="navbar-nav me-auto mb-2 mb-lg-0">
						<li className="nav-item">
							<a className="nav-link active" aria-current="page" href="https://gitbrent.github.io/PptxGenJS/demo/react/">
								Vite+React Demo Home
							</a>
						</li>
					</ul>
					<div className="hstack gap-1">
						<button type="button" className="btn btn-primary" title="Releases" onClick={() => "window.open('https://github.com/gitbrent/PptxGenJS/releases')"}>
							<i className="bi bi-box-arrow-up-right me-2"></i>Latest Release
						</button>
						<button type="button" className="btn btn-primary" title="Docs" onClick={() => "window.open('https://gitbrent.github.io/PptxGenJS/docs/installation/')"}>
							<i className="bi bi-box-arrow-up-right me-2"></i>Docs
						</button>
						<div className="vr my-2 mx-2"></div>
						<button type="button" className="btn btn-primary" title="@gitbrent@fosstodon.org" onClick={() => "window.open('https://fosstodon.org/@gitbrent')"}>
							<i className="bi bi-mastodon"></i>
						</button>
						<button type="button" className="btn btn-primary" title="GitHub" onClick={() => "window.open('https://gitbrent.github.io/PptxGenJS')"}>
							<i className="bi bi-github"></i>
						</button>
					</div>
				</div>
			</div>
		</nav>
	}

	const htmlMain = () => {
		return <main className="container my-5">
			<div className="card">
				<div className="card-header">
					<h1 className="display-4">Module Demo</h1>
					<div className="lead text-primary-emphasis">
						Sample React+TypeScript+Vite application demonstrating the PptxGenJS library as a module.
					</div>
				</div>
				<div className="card-body">
					<h5 className="text-info">Demo Code (.tsx)</h5>
					<pre className="bg-black mt-3">
						<code className="language-javascript" style={{ fontSize: "0.75rem" }}>{demoCode}</code>
					</pre>
					<table id="html2ppt" className="table table-dark d-none">
						<thead className="table-dark">
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
				<div className="card-footer p-3">
					<div className="row row-cols-1 row-cols-md-3 g-4">
						<div className="col">
							<button type="button" className="btn btn-success w-100" onClick={() => runDemo()}>
								<h2>Run Test 1</h2>
								Demo Code
							</button>
						</div>
						<div className="col">
							<button type="button" className="btn btn-primary w-100" onClick={() => testMainMethods()}>
								<h2>Run Test 2</h2>
								Misc Objects
							</button>
						</div>
						<div className="col">
							<button type="button" className="btn btn-primary w-100" onClick={() => testTableMethod()}>
								<h2>Run Test 3</h2>
								Table-to-Slides
							</button>
						</div>
					</div>
				</div>
			</div>
		</main>
	}

	return (
		<section>
			{htmlNav()}
			{htmlMain()}
		</section>
	);
}

export default App
