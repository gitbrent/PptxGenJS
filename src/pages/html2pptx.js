import React from "react";
import Layout from "@theme/Layout";
import Gist from "react-gist";
//import "../css/bootstrap-yeti.css";
import "../css/purged.css";

export default () => {
	const Header = () => {
		return (
			<div className="card useTheme mb-5">
				<div className="card-header h4 bg-primary text-white">About</div>
				<div className="card-body">
					<p>Create a presentation from an HTML table with a single line of code.</p>
					<p>
						The <code>tableToSlides</code> method clones the table including CSS style and creates slides as needed (auto-paging).
					</p>
					<div className="row align-items-center my-3">
						<div className="col-auto">
							<img
								className="d-none d-md-none d-lg-block border border-light"
								alt="input: html table"
								src="/PptxGenJS/img/ex-html-to-powerpoint-1.png"
								height="500"
							/>
							<img
								className="d-none d-md-block d-lg-none border border-light"
								alt="input: html table"
								src="/PptxGenJS/img/ex-html-to-powerpoint-1.png"
								height="300"
							/>
							<img
								className="d-block d-md-none d-lg-none border border-light"
								alt="input: html table"
								src="/PptxGenJS/img/ex-html-to-powerpoint-1.png"
								height="200"
							/>
						</div>
						<div className="col-auto col-md text-center">
							<h1 className="mb-0">→</h1>
						</div>
						<div className="col-auto">
							<img
								className="d-none d-md-none d-lg-block border border-light"
								alt="output: powerpoint slides"
								src="/PptxGenJS/img/ex-html-to-powerpoint-2.png"
								height="500"
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
					<p className="mb-0">
						Refer to the HTML-to-PowerPoint docs for a complete list of options like repeating header, start location on subsequnet slides, character and line weight,
						etc.
					</p>
				</div>
				<div className="card-footer">
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

	const Body = () => {
		// NOTE: 20210407: cant use `row-cols-12` yet as docusaurus core bootstrap is ruining `rows` style, use "col-12" etc on cols to supercede for now
		return (
			<div className="card useTheme mb-5">
				<div className="card-header h4 bg-primary text-white">Code</div>
				<div className="card-body">
					<p>These 3 lines of code were all it took to produce the slides shown above!</p>
					<Gist id="494850b6775dd5c8ce314672a1846208" />
				</div>
				<div className="card-footer">
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

	const Footer = () => {
		return (
			<div className="card useTheme mb-5">
				<div className="card-header h4 bg-primary text-white">Demo</div>
				<div className="card-body">
					<p>Try it for yourself! The demo below has an interactive table plus various options to test drive.</p>
					<div className="text-center">
						<img alt="HTML Table" src="/PptxGenJS/img/ex-html-to-powerpoint-3.png" height="500" className="border border-light" />
					</div>
				</div>
				<div className="card-footer">
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
				<Header />
				<Body />
				<Footer />
			</div>
		</Layout>
	);
};
