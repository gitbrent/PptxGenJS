import React from "react";
import Layout from "@theme/Layout";
import useThemeContext from "@theme/hooks/useThemeContext";
//import "../css/bootstrap.css";
//import "../css/bootstrap-yeti.css";
import "../css/purged.css";

export default () => {
	const Header = () => {
		const { isDarkTheme, setLightTheme, setDarkTheme } = useThemeContext();

		return (
			<header className={isDarkTheme ? "header-dark" : "header-light"}>
				<div className="container">
					<div className="row justify-content-center">
						<div className="col-auto">
							<div className="my-5">
								<h1 className={`display-1 mb-3 ${isDarkTheme ? "text-light" : "text-primary"}`}>PptxGenJS</h1>
								<h3 className={`fw-light mb-4 ${isDarkTheme ? "text-white-50" : "text-black-50"}`}>Create PowerPoint presentations with JavaScript</h3>
								<h6 className={`fw-light mb-3 ${isDarkTheme ? "text-white-50" : "text-black-50"}`}>
									The most popular powerpoint+js library on npm with nearly 1,000 stars on GitHub
								</h6>
								<div className="row row-cols-1 row-cols-md-2 g-4 my-0">
									<div className="col-12 col-md-4">
										<button
											type="button"
											aria-label="Get Started"
											className={`w-100 fw-bold btn py-3 ${isDarkTheme ? "btn-primary" : "btn-primary"}`}
											onClick={() => (window.location.href = "/PptxGenJS/docs/quick-start")}
										>
											Get Started
										</button>
									</div>
									<div className="col-12 col-md-4">
										<button
											type="button"
											aria-label="View Demos"
											className={`w-100 fw-bold btn py-3 ${isDarkTheme ? "btn-outline-primary" : "btn-outline-primary"}`}
											onClick={() => (window.location.href = "/PptxGenJS/pptxdemos")}
										>
											Demos
										</button>
									</div>
									<div className="col-12 col-md-4">
										<button
											type="button"
											aria-label="Learn about HTML to PowerPoint"
											className={`w-100 fw-bold btn fw-bold py-3 ${isDarkTheme ? "btn-outline-primary" : "btn-outline-primary"}`}
											onClick={() => (window.location.href = "/PptxGenJS/html2pptx")}
										>
											HTML to PPTX
										</button>
									</div>
								</div>
							</div>
						</div>
					</div>
				</div>
			</header>
		);
	};

	const Body = () => {
		const { isDarkTheme, setLightTheme, setDarkTheme } = useThemeContext();
		return (
			<main className={`py-5 ${isDarkTheme ? "body-dark" : "bg-light"}`}>
				<div className="container">
					<div className="row g-5 mb-0">
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "border-0" : ""}`}>
								<div className={`card-body border-top border-2 border-primary p-4 border-0 ${isDarkTheme ? "bg-black text-white" : "bg-white"}`}>
									<h4 className="text-primary mb-4">Works Everywhere</h4>
									<ul className={`mb-0 ${isDarkTheme ? "text-white-50" : "text-black-50"}`}>
										<li>Every modern desktop and mobile browser is supported</li>
										<li>Integrates with Node, Angular, React and Electron</li>
										<li>Compatible with PowerPoint, Keynote, and more</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "border-0" : ""}`}>
								<div className={`card-body border-top border-2 border-primary p-4 border-0 ${isDarkTheme ? "bg-black text-white" : "bg-white"}`}>
									<h4 className="text-primary mb-4">Full Featured</h4>
									<ul className={`mb-0 ${isDarkTheme ? "text-white-50" : "text-black-50"}`}>
										<li>All major object types are available (charts, shapes, tables, etc.)</li>
										<li>Master Slides for academic/corporate branding</li>
										<li>SVG images, animated gifs, YouTube videos, RTL text, and Asian fonts</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "border-0" : ""}`}>
								<div className={`card-body border-top border-2 border-primary p-4 border-0 ${isDarkTheme ? "bg-black text-white" : "bg-white"}`}>
									<h4 className="text-primary mb-4">Simple And Powerful</h4>
									<ul className={`mb-0 ${isDarkTheme ? "text-white-50" : "text-black-50"}`}>
										<li>The absolute easiest PowerPoint library to use</li>
										<li>Learn as you code will full typescript definitions included</li>
										<li>Tons of demo code comes included (over 70 slides of features)</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "border-0" : ""}`}>
								<div className={`card-body border-top border-2 border-primary p-4 border-0 ${isDarkTheme ? "bg-black text-white" : "bg-white"}`}>
									<h4 className="text-primary mb-4">Export Your Way</h4>
									<ul className={`mb-0 ${isDarkTheme ? "text-white-50" : "text-black-50"}`}>
										<li>Exports files direct to client browsers with proper MIME-type</li>
										<li>Other export formats available: base64, blob, stream, etc.</li>
										<li>Presentation compression options and more</li>
									</ul>
								</div>
							</div>
						</div>
					</div>
				</div>
			</main>
		);
	};

	return (
		<Layout title="Home">
			<Header />
			<Body />
		</Layout>
	);
};
