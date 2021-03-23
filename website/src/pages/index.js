import React from "react";
import Layout from "@theme/Layout";
import useThemeContext from "@theme/hooks/useThemeContext";
import "../css/bootstrap.min.css";
//import "../css/purged.min.css";

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
											onClick={() => (window.location = "./demo")}
										>
											Online Demo
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
					<div className="row row-cols-1 row-cols-md-2 g-5">
						<div className="col">
							<div className={`card ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Awesome</h5>
								</div>
								<div className="card-body">&bull;&nbsp;WOW!</div>
							</div>
						</div>
						<div className="col">
							<div className={`card ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Awesome</h5>
								</div>
								<div className="card-body">&bull;&nbsp;WOW!</div>
							</div>
						</div>
						<div className="col">
							<div className={`card ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Awesome</h5>
								</div>
								<div className="card-body">&bull;&nbsp;WOW!</div>
							</div>
						</div>
						<div className="col">
							<div className={`card ${isDarkTheme ? "bg-secondary text-white" : "bg-light"}`}>
								<div className="card-header">
									<h5 className="fw-light mb-0">Awesome</h5>
								</div>
								<div className="card-body">&bull;&nbsp;WOW!</div>
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
