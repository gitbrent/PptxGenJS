import React from "react";
import Layout from "@theme/Layout";
//import "../css/bootstrap-yeti.css";
import "../css/purged.css";

export default () => {
	return (
		<Layout title="Demos">
			<div className="container my-4">
				<h1 className="mb-4">Demo</h1>
				<div className="alert alert-info">We've moved! (You arrived here via a pre-3.7.0 README file)</div>
				<p>
					Please visit: <a href="https://gitbrent.github.io/PptxGenJS/pptxdemos">https://gitbrent.github.io/PptxGenJS/pptxdemos</a>
				</p>
			</div>
		</Layout>
	);
};
