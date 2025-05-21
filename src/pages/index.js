import React from "react";
import Layout from "@theme/Layout";

export default () => {
	const Header = () => {
		return (
			<header id="indexHeader">
				<div className="row justify-content-center">
					<div className="col-auto">
						<div className="my-5">
							<h1 className="display-1 mb-3">PptxGenJS</h1>
							<h3 className="mb-4">
								Build PowerPoint presentations with JavaScript!<br />
								Works with Node, React, web browsers, and more.
							</h3>
							<h6 className="mb-4">
								The most popular powerpoint+js library on npm with 3,500 stars on GitHub
							</h6>
							<div className="row row-cols row-cols-4-md mt-4">
								<div className="col">
									<button
										type="button"
										aria-label="Get Started"
										className="w-100 fw-bold btn btn-lg py-3 btn-success"
										onClick={() => (window.location.href = "/PptxGenJS/docs/quick-start/")}
									>
										Get Started
									</button>
								</div>
								<div className="col">
									<button
										type="button"
										aria-label="View Demos"
										className="w-100 fw-bold btn btn-lg py-3 btn-primary"
										onClick={() => (window.location.href = "/PptxGenJS/demos/")}
									>
										Demos
									</button>
								</div>
								<div className="col">
									<button
										type="button"
										aria-label="Table-to-Slides Feature"
										className="w-100 fw-bold btn btn-lg py-3 btn-primary"
										onClick={() => (window.location.href = "/PptxGenJS/html2pptx/")}
									>
										HTML to PPTX
									</button>
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
			<main className="useTheme pb-4">
				<div className="row row-cols-2 row-cols-md-2 g-4 mb-0">
					<div className="col">
						<div className="card h-100">
							<div className="card-body border-top border-primary border-5">
								<h4>Works Everywhere</h4>
								<ul className="mb-0">
									<li>Supports every major modern browser - desktop and mobile</li>
									<li>Seamlessly integrates with Node.js, React, Angular, Vite, and Electron</li>
									<li>Compatible with PowerPoint, Keynote, LibreOffice, and other apps</li>
								</ul>
							</div>
						</div>
					</div>
					<div className="col">
						<div className="card h-100">
							<div className="card-body border-top border-primary border-5">
								<h4>Full Featured</h4>
								<ul className="mb-0">
									<li>Create all major slide objects: text, tables, shapes, images, charts, and more</li>
									<li>Define custom Slide Masters for consistent academic or corporate branding</li>
									<li>Supports SVGs, animated GIFs, YouTube embeds, RTL text, and Asian fonts</li>
								</ul>
							</div>
						</div>
					</div>
					<div className="col">
						<div className="card h-100">
							<div className="card-body border-top border-primary border-5">
								<h4>Simple &amp; Powerful</h4>
								<ul className="mb-0">
									<li>Ridiculously easy to use - create a presentation in 4 lines of code</li>
									<li>Full TypeScript definitions for autocomplete and inline documentation</li>
									<li>Includes 75+ demo slides covering every feature and usage pattern</li>
								</ul>
							</div>
						</div>
					</div>
					<div className="col">
						<div className="card h-100">
							<div className="card-body border-top border-primary border-5">
								<h4>Export Your Way</h4>
								<ul className="mb-0">
									<li>Instantly download .pptx files from the browser with proper MIME handling</li>
									<li>Export as base64, Blob, Buffer, or Node stream</li>
									<li>Supports compression and advanced output options for production use</li>
								</ul>
							</div>
						</div>
					</div>
				</div>
			</main>
		);
	};

	return (
		<Layout title="Home">
			<div id="home">
				<div className="container mb-5">
					<Header />
					<Body />
				</div>
			</div>
		</Layout>
	);
};
