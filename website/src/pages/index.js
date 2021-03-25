import React from "react";
import Layout from "@theme/Layout";
import useThemeContext from "@theme/hooks/useThemeContext";
//import "../css/bootstrap.min.css";
import "../css/purged.min.css";

export default () => {
	const Header = () => {
		const { isDarkTheme, setLightTheme, setDarkTheme } = useThemeContext();
		return (
			<header style={{ backgroundColor: isDarkTheme ? "var(--bs-gray-dark)" : "var(--bs-light)" }}>
				<div className="container">
					<div className="row justify-content-center">
						<div className="col-auto">
							<div className="my-5">
								<h1 className={`display-1 mb-3 ${isDarkTheme ? "text-light" : ""}`}>PptxGenJS</h1>
								<h3 className={`fw-light mb-4 ${isDarkTheme ? "text-muted" : "text-muted"}`}>Create PowerPoint presentations with JavaScript</h3>
								<h6 className="text-muted fw-light mb-4">The most popular powerpoint+js library on npm with nearly 1,000 stars on GitHub</h6>
								<div className="row row-cols-1 row-cols-md-2 g-4 my-0">
									<div className="col">
										<button
											type="button"
											className={`w-100 btn ${isDarkTheme ? "btn-outline-light" : "btn-outline-dark"}`}
											onClick={() => (window.location = "./docs/quick-start")}
										>
											Get Started
										</button>
									</div>
									<div className="col">
										<button
											type="button"
											className={`w-100 btn ${isDarkTheme ? "btn-outline-light" : "btn-outline-dark"}`}
											onClick={() => (window.location = "./docs/demos")}
										>
											Online Demos
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
			<main className={`py-5 ${isDarkTheme ? "bg-dark" : "bg-white"}`}>
				<div className="container">
					<div className="row g-5 mb-0">
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Works Everwhere</h5>
								</div>
								<div className={`card-body ${isDarkTheme ? "bg-dark" : "bg-white"}`}>
									<ul className="mb-0">
										<li>Every modern desktop and mobile browser is supported</li>
										<li>Integrates with Node, Angular, React and Electron</li>
										<li>IE11 support available via polyfill</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Full Featured</h5>
								</div>
								<div className={`card-body ${isDarkTheme ? "bg-dark" : "bg-white"}`}>
									<ul className="mb-0">
										<li>All major object types are available (charts, shapes, tables, etc.)</li>
										<li>Master Slides for academic/corporate branding</li>
										<li>SVG images, animated gifs, YouTube videos, RTL text, and Asian fonts</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Simple And Powerful</h5>
								</div>
								<div className={`card-body ${isDarkTheme ? "bg-dark" : "bg-white"}`}>
									<ul className="mb-0">
										<li>The absolute easiest PowerPoint library to use</li>
										<li>Learn as you code will full typescript definitions included</li>
										<li>Tons of demo code comes included (almost 80 slides of features)</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Export Your Way</h5>
								</div>
								<div className={`card-body ${isDarkTheme ? "bg-dark" : "bg-white"}`}>
									<ul className="mb-0">
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
