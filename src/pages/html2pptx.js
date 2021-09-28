import React from "react";
import Layout from "@theme/Layout";
import Gist from "react-gist";
//import "../css/bootstrap-yeti.css";
import "../css/purged.css";

export default () => {
	const Header = () => {
		return (
			<section className="bgTheme p-4">
				<h3 className="mb-3">About HTML-to-PPTX</h3>
				<p>
					The <code>tableToSlides</code> method generates a presentation from an HTML table element id.
				</p>
				<ul>
					<li>Many options are available including repeating header, start location on subsequent slides, character and line weight</li>
					<li>Additional slides are automatically created as needed (auto-paging)</li>
					<li>The table's style (CSS) is copied into the PowerPoint table</li>
				</ul>
				<div className="d-none d-md-flex row align-items-center justify-content-center my-3">
					<div className="col-auto">
						<img className="d-none d-md-none d-lg-block border border-light" alt="input: html table" src="/PptxGenJS/img/ex-html-to-powerpoint-1.png" height="400" />
						<img className="d-none d-md-block d-lg-none border border-light" alt="input: html table" src="/PptxGenJS/img/ex-html-to-powerpoint-1.png" height="300" />
						<img className="d-block d-md-none d-lg-none border border-light" alt="input: html table" src="/PptxGenJS/img/ex-html-to-powerpoint-1.png" height="200" />
					</div>
					<div className="col-auto">
						<h1 className="mb-0">→</h1>
					</div>
					<div className="col-auto">
						<img
							className="d-none d-md-none d-lg-block border border-light"
							alt="output: powerpoint slides"
							src="/PptxGenJS/img/ex-html-to-powerpoint-2.png"
							height="400"
						/>
						<img
							className="d-none d-md-block d-lg-none border border-light"
							alt="output: powerpoint slides"
							src="/PptxGenJS/img/ex-html-to-powerpoint-2.png"
							height="300"
						/>
						<img
							className="d-block d-md-none d-lg-none border border-light"
							alt="output: powerpoint slides"
							src="/PptxGenJS/img/ex-html-to-powerpoint-2.png"
							height="200"
						/>
					</div>
				</div>
			</section>
		);
	};

	const SampleCode = () => {
		// NOTE: 20210407: cant use `row-cols-12` yet as docusaurus core bootstrap is ruining `rows` style, use "col-12" etc on cols to supercede for now
		return (
			<div className="card useTheme h-100">
				<div className="card-body">
					<h3 className="mb-3">Sample Code</h3>
					<p>Reproduce a table in as little as 3 lines of code.</p>
					<Gist id="494850b6775dd5c8ce314672a1846208" />
				</div>
				<div className="card-footer text-center">
					<button
						type="button"
						aria-label="documentation"
						className="btn btn-outline-primary px-5"
						onClick={() => (window.location.href = "/PptxGenJS/docs/html-to-powerpoint")}
					>
						HTML to PowerPoint Docs
					</button>
				</div>
			</div>
		);
	};

	const LiveDemo = () => {
		return (
			<div className="card useTheme h-100">
				<div className="card-body">
					<h3 className="mb-3">Live Demo</h3>
					<p>Try the html-to-pptx feature out for yourself.</p>
					<div className="text-center">
						<img alt="HTML Table" src="/PptxGenJS/img/ex-html-to-powerpoint-3.png" className="border border-light" />
					</div>
				</div>
				<div className="card-footer text-center">
					<button
						type="button"
						aria-label="demo"
						className="btn btn-outline-primary px-5"
						onClick={() => (window.location.href = "/PptxGenJS/demo/index.html#html2pptx")}
					>
						HTML to PowerPoint Demo
					</button>
				</div>
			</div>
		);
	};

	return (
		<Layout title="HTML-to-PowerPoint">
			<div className="container my-4">
				<h1 className="mb-4">HTML to PowerPoint</h1>
				<div className="row g-5">
					<div className="col-12">
						<Header />
					</div>
					<div className="col">
						<SampleCode />
					</div>
					<div className="col">
						<LiveDemo />
					</div>
				</div>
			</div>
		</Layout>
	);
};
