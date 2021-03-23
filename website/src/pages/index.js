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
								<h3 className={`fw-light mb-4 ${isDarkTheme ? "text-muted" : "text-muted"}`}>Create JavaScript PowerPoint Presentations</h3>
								<h6 className="text-muted fw-light mb-4">The most forked javascript+powerpoint project on GitHub with nearly 1,000 stars</h6>
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
									<h5 className="fw-light mb-0">Works Everywhere</h5>
								</div>
								<div className={`card-body ${isDarkTheme ? "bg-dark" : "bg-white"}`}>
									<ul className="mb-0">
										<li>Every modern desktop and mobile browser is supported</li>
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
										<li>All major object types are available (charts, tables, etc.)</li>
										<li>Master Slides for corporate branding are supported</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Simple To Use</h5>
								</div>
								<div className={`card-body ${isDarkTheme ? "bg-dark" : "bg-white"}`}>
									<ul className="mb-0">
										<li>The easiest PowerPoint library to use</li>
										<li>Create a presentation using just a few lines of code</li>
										<li>Learn as you code will full typescript definitions included</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className={`card h-100 ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Modern</h5>
								</div>
								<div className={`card-body ${isDarkTheme ? "bg-dark" : "bg-white"}`}>
									<ul className="mb-0">
										<li>Integrates with client browsers, Node, Angular, React and Electron</li>
										<li>Export files direct to browser, blob, stream and more</li>
										<li>Presentation compression supported</li>
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
