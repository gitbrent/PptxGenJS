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
						<img alt="HTML Table" src="img/ex-html-to-powerpoint-1.png" height="500" className="border border-light" />
					</div>
					<div className="col text-center">
						<h1 className="mb-0">→</h1>
					</div>
					<div className="col-auto">
						<img alt="PowerPoint with HTML Table" src="img/ex-html-to-powerpoint-2.png" height="500" />
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
					The <code>tableToSlides</code> method clones the table including CSS style and creates slides as needed (auto-paging).
				</p>
				<p>
					View the <a href="docs/html-to-powerpoint">HTML-to-PowerPoint docs</a> for a complete list of options like repeating header, start location on subsequnet
					slides, character and line weight, etc.
				</p>
				<Gist id="494850b6775dd5c8ce314672a1846208" />
			</section>
		);
	};

	const Footer = () => {
		const { isDarkTheme, setLightTheme, setDarkTheme } = useThemeContext();

		return (
			<section className={`mb-5 p-4 ${isDarkTheme ? "bg-black" : "bg-white"}`}>
				<h4>Live Demo</h4>
				<p>
					The <a href="demo/index.html#html2pptx">complete library demo</a> has an interactive table and options you can interact with.
				</p>
				<div className="text-center">
					<img alt="HTML Table" src="img/ex-html-to-powerpoint-3.png" height="500" className="border border-light" />
				</div>
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
