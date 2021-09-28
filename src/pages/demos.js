import React from "react";
import Layout from "@theme/Layout";
import pptxgen from "pptxgenjs";
import Gist from "react-gist";
//import "../css/bootstrap-yeti.css";
import "../css/purged.css";

export default () => {
	const gistRef = React.useRef(null);

	/**
	 * @src https://gist.github.com/gitbrent/84acbcaab54be0eba83f5206ef6ddd95.js
	 */
	function doLiveDemo() {
		let pptx = new pptxgen();
		let slide = pptx.addSlide();

		// FUTURE: read actual code from gist
		/*
			console.log(gistRef);
			let gist = gistRef.current
			.querySelector("#file-pptxgenjs_demo-js")
			.querySelectorAll(".blob-code")
			.forEach((item) => console.log(item));
		*/
		slide.addText("BONJOUR - CIAO - GUTEN TAG - HELLO - HOLA - NAMASTE - 你好", {
			x: 0,
			y: 1,
			w: "100%",
			h: 2,
			align: "center",
			color: "0088CC",
			fill: "F1F1F1",
			fontSize: 24,
		});

		pptx.writeFile({ fileName: "PptxGenJS-Demo" });
	}

	const LiveDemo = () => {
		return (
			<section className="bgTheme mb-5 p-4">
				<h3 className="mb-3">Live Demo</h3>
				<p>Creating a presentation really is this easy! Output can also be produced as base64, blob, stream and more.</p>
				<Gist id="84acbcaab54be0eba83f5206ef6ddd95" ref={gistRef} />
				<div className="row justify-content-between">
					<div className="col-auto">
						<button type="button" className="btn btn-primary" onClick={() => doLiveDemo()}>
							Run Live Demo
						</button>
					</div>
					<div className="col-auto">
						<button type="button" className="btn btn-primary" onClick={() => window.open("https://jsfiddle.net/gitbrent/L1uctxm0/", "_blank")}>
							Code Live Using jsFiddle...
						</button>
					</div>
				</div>
			</section>
		);
	};

	const BothDemos = () => {
		// NOTE: 20210407: cant use `row-cols-12` yet as docusaurus core bootstrap is ruining `rows` style, use "col-12" etc on cols to supercede for now
		return (
			<section>
				<div className="row g-5 mb-5">
					<div className="col-12 col-md-6">
						<div className="bgTheme h-100 p-4">
							<h3 className="mb-3">Complete Library Demo</h3>
							<p>
								Function demos for every feature are available in the <a href="https://github.com/gitbrent/PptxGenJS/tree/master/demos/browser">browser demo</a>,
								which is hosted online below. Over 70 slides worth of various PowerPoint objects can be produced.
							</p>
							<button type="button" className="btn btn-primary" onClick={() => window.open("/PptxGenJS/demo/browser/index.html", "_blank")}>
								Complete Library Demo
							</button>
						</div>
					</div>
					<div className="col-12 col-md-6">
						<div className="bgTheme h-100 p-4">
							<h4>React App Demo</h4>
							<p>
								There is a complete react application available in the <a href="https://github.com/gitbrent/PptxGenJS/tree/master/demos/react-demo">demos/react</a>{" "}
								folder. The latest build can be run as a demo below.
							</p>
							<button type="button" className="btn btn-primary" onClick={() => window.open("/PptxGenJS/demo/react/index.html", "_blank")}>
								React App Demo
							</button>
						</div>
					</div>
				</div>
			</section>
		);
	};

	const BigImage = () => {
		return (
			<section className="bgTheme mb-5 p-4">
				<h3 className="mb-3">Demo Slides</h3>
				<p>
					The complete library demo above has 70+ slides worth of demo code that you can use to get started.
					<br />
					All demo code is divided into modules on github under <a href="https://github.com/gitbrent/PptxGenJS/tree/master/demos/modules">demos/modules</a>.
				</p>
				<img alt="PptxGenJS PowerPoint Demo Slides" src="/PptxGenJS/img/readme_banner.png" />
			</section>
		);
	};

	return (
		<Layout title="Demos">
			<div className="container my-4">
				<h1 className="mb-4">Demos</h1>
				<LiveDemo />
				<BothDemos />
				<BigImage />
			</div>
		</Layout>
	);
};
