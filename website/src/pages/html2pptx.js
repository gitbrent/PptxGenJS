import React from "react";
import Layout from "@theme/Layout";
import useThemeContext from "@theme/hooks/useThemeContext";
import Gist from "react-gist";
//import "../css/bootstrap-yeti.css";
import "../css/purged.css";

export default () => {
	const Header = () => {
		const { isDarkTheme, setLightTheme, setDarkTheme } = useThemeContext();

		return (
			<section className={`mb-5 p-4 ${isDarkTheme ? "bg-black" : "bg-white"}`}>
				<h4>Table to Slides Feature</h4>
				<p>Create a presentation from an HTML table with a single line of code. Creates slides as needed (auto-paging).</p>
				<div className="row align-items-center">
					<div className="col-auto">
						<img className="d-none d-md-none d-lg-block border border-light" alt="input: html table" src="/PptxGenJS/img/ex-html-to-powerpoint-1.png" height="500" />
						<img className="d-none d-md-block d-lg-none border border-light" alt="input: html table" src="/PptxGenJS/img/ex-html-to-powerpoint-1.png" height="300" />
						<img className="d-block d-md-none d-lg-none border border-light" alt="input: html table" src="/PptxGenJS/img/ex-html-to-powerpoint-1.png" height="200" />
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
			</section>
		);
	};

	const Body = () => {
		const { isDarkTheme, setLightTheme, setDarkTheme } = useThemeContext();

		// NOTE: 20210407: cant use `row-cols-12` yet as docusaurus core bootstrap is ruining `rows` style, use "col-12" etc on cols to supercede for now
		return (
			<section className={`mb-5 p-4 ${isDarkTheme ? "bg-black" : "bg-white"}`}>
				<h4>Demo Code</h4>
				<p>
					The <code>tableToSlides</code> method clones the table including CSS style and creates slides as needed (auto-paging). These 3 lines of code were all it took to
					produce the slides shown above!
				</p>
				<p>View the HTML-to-PowerPoint docs for a complete list of options like repeating header, start location on subsequnet slides, character and line weight, etc.</p>
				<Gist id="494850b6775dd5c8ce314672a1846208" />
				<button
					type="button"
					aria-label="HTML to PowerPoint Documentation"
					className={`w-100 mt-4 btn ${isDarkTheme ? "btn-outline-primary" : "btn-outline-primary"}`}
					onClick={() => (window.location.href = "/PptxGenJS/docs/html-to-powerpoint")}
				>
					HTML to PowerPoint Docs
				</button>
			</section>
		);
	};

	const Footer = () => {
		const { isDarkTheme, setLightTheme, setDarkTheme } = useThemeContext();

		return (
			<section className={`mb-5 p-4 ${isDarkTheme ? "bg-black" : "bg-white"}`}>
				<h4>Live Demo</h4>
				<p>Try it for yourself! The demo below has an interactive table plus various options to test drive.</p>
				<div className="text-center">
					<img alt="HTML Table" src="/PptxGenJS/img/ex-html-to-powerpoint-3.png" height="500" className="border border-light" />
				</div>
				<button
					type="button"
					aria-label="HTML to PowerPoint Demo"
					className={`w-100 mt-4 btn ${isDarkTheme ? "btn-outline-primary" : "btn-outline-primary"}`}
					onClick={() => (window.location.href = "/PptxGenJS/demo/index.html#html2pptx")}
				>
					HTML to PowerPoint Demo
				</button>
			</section>
		);
	};

	return (
		<Layout title="PptxGenJS Demos">
			<div className="container my-4">
				<h1 className="mb-4">HTML to PowerPoint</h1>
				<Header />
				<Body />
				<Footer />
			</div>
		</Layout>
	);
};
