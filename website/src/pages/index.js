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
								<h1 className={`display-1 mb-4 ${isDarkTheme ? "text-light" : ""}`}>PptxGenJS</h1>
								<h3 className={`fw-light mb-2 ${isDarkTheme ? "text-muted" : "text-muted"}`}>Create JavaScript PowerPoint Presentations</h3>
								<h6 className="text-muted fw-light mb-4">The fastest growing, second most starred powerpoint library on GitHub</h6>
								<div className="text-muted fw-light mb-4 fw-light">The fastest growing, 2nd most starred powerpoint library on GitHub</div>
								<div className="row row-cols-1 row-cols-md-2 g-4">
									<div className="col">
										<button type="button" className={`w-100 btn ${isDarkTheme ? "btn-outline-light" : "btn-outline-dark"}`}>
											Get Started
										</button>
									</div>
									<div className="col">
										<button type="button" className={`w-100 btn ${isDarkTheme ? "btn-outline-light" : "btn-outline-dark"}`}>
											Features
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

	return (
		<Layout title="pptxgenjs home">
			<div>
				<Header />
				<main className="container">
					<div className="row row-cols-2 row-cols-md-4 g-4">
						<div className="col">COL1</div>
						<div className="col">COL2</div>
						<div className="col">COL3</div>
						<div className="col">COL4</div>
					</div>
				</main>
			</div>
		</Layout>
	);
};
