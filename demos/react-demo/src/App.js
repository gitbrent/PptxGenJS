import React from "react";
import logo from "./logo.svg";
import "./App.css";
import pptxgen from "./pptxgen.es.js"; // LOCAL DEV TESTING src=`PptxGenJS/dist`
//import pptxgen from "pptxgenjs"; // react-app webpack will use package.json `"module": "dist/pptxgen.es.js"` value

function App() {
	const demoCode = `import pptxgen from "pptxgenjs";
let pptx = new pptxgen();

let slide = pptx.addSlide();
slide.addText("React Demo!", { x:1, y:1, w:'80%', h:1, fontSize:36, fill:'eeeeee', align:'center' });

pptx.writeFile("react-demo.pptx");`;

	function runDemo() {
		let pptx = new pptxgen();
		let slide = pptx.addSlide();
		slide.addText("React Demo!", { x: 1, y: 1, w: "80%", h: 1, fontSize: 36, fill: "eeeeee", align: "center" });
		pptx.writeFile("react-demo.pptx");

		console.log(`pptx.version = ${pptx.version}`);
	}

	return (
		<div>
			<nav className="navbar navbar-expand-lg navbar-dark bg-primary">
				<a className="navbar-brand" href="https://gitbrent.github.io/PptxGenJS/">
					<img src={logo} width="30" height="30" className="d-inline-block align-top mr-2" alt="" />
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
							onClick={ev => {
								window.open("https://gitbrent.github.io/PptxGenJS/demo/", true);
							}}
						>
							Demo Page
						</button>
						<button
							type="button"
							className="btn btn-outline-info mx-3 my-2 my-sm-0"
							onClick={ev => {
								window.open("https://github.com/gitbrent/PptxGenJS", true);
							}}
						>
							GitHub Project
						</button>

						<button
							type="button"
							className="btn btn-outline-info mx-3 my-2 my-sm-0"
							onClick={ev => {
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

					<h5 className="text-info">Demo Code</h5>
					<pre className="my-4">
						<code className="language-javascript">{demoCode}</code>
					</pre>

					<button type="button" className="btn btn-success w-25" onClick={ev => runDemo()}>
						Run Demo
					</button>
				</div>
			</main>
		</div>
	);
}

export default App;
