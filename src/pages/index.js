import React from "react";
import Layout from "@theme/Layout";
//import "../css/bootstrap.css";
//import "../css/bootstrap-yeti.css";
import "../css/purged.css";

export default () => {
	const Header = () => {
		return (
			<header id="indexHeader">
				<div className="container">
					<div className="row justify-content-center">
						<div className="col-auto">
							<div className="my-5">
								<h1 className="display-1 mb-3">PptxGenJS</h1>
								<h3 className="fw-light mb-4">Create PowerPoint presentations with JavaScript</h3>
								<h6 className="fw-light mb-3">The most popular powerpoint+js library on npm with over 1,500 stars on GitHub</h6>
								<div className="row row-cols-1 row-cols-md-2 g-4 my-0">
									<div className="col-12 col-md-4">
										<button
											type="button"
											aria-label="Get Started"
											className="w-100 fw-bold btn py-3 btn-primary"
											onClick={() => (window.location.href = "/PptxGenJS/docs/quick-start/")}
										>
											Get Started
										</button>
									</div>
									<div className="col-12 col-md-4">
										<button
											type="button"
											aria-label="View Demos"
											className="w-100 fw-bold btn py-3 btn-outline-primary"
											onClick={() => (window.location.href = "/PptxGenJS/demos/")}
										>
											Demos
										</button>
									</div>
									<div className="col-12 col-md-4">
										<button
											type="button"
											aria-label="Learn about HTML to PowerPoint"
											className="w-100 fw-bold btn py-3 btn-outline-primary"
											onClick={() => (window.location.href = "/PptxGenJS/html2pptx/")}
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
		return (
			<main className="useTheme py-5">
				<div className="container">
					<div className="row g-5 mb-0">
						<div className="col-12 col-md-6">
							<div className="card h-100 border-0">
								<div className="card-body border-top border-2 border-primary p-4 border-0">
									<h4 className="text-primary mb-4">Works Everywhere</h4>
									<ul className="mb-0">
										<li>Every modern desktop and mobile browser is supported</li>
										<li>Integrates with Node, Angular, React and Electron</li>
										<li>Compatible with Microsoft PowerPoint, Apple Keynote, and many others</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className="card h-100 border-0">
								<div className="card-body border-top border-2 border-primary p-4 border-0">
									<h4 className="text-primary mb-4">Full Featured</h4>
									<ul className="mb-0">
										<li>All major objects are available (charts, shapes, tables, etc)</li>
										<li>Master Slide support for academic/corporate branding</li>
										<li>SVG images, animated gifs, YouTube videos, RTL text, and Asian fonts</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className="card h-100 border-0">
								<div className="card-body border-top border-2 border-primary p-4 border-0">
									<h4 className="text-primary mb-4">Simple And Powerful</h4>
									<ul className="mb-0">
										<li>The absolute easiest PowerPoint library to use</li>
										<li>Learn as you code using the built-in typescript definitions</li>
										<li>Tons of sample code comes included (75+ slides of demos)</li>
									</ul>
								</div>
							</div>
						</div>
						<div className="col-12 col-md-6">
							<div className="card h-100 border-0">
								<div className="card-body border-top border-2 border-primary p-4 border-0">
									<h4 className="text-primary mb-4">Export Your Way</h4>
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
