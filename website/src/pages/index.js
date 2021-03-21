import Layout from "@theme/Layout";
import React from "react";
import "../css/bootstrap.min.css";

export default () => {
	return (
		<Layout title="pptxgenjs home">
			<div className="container">
				<header className="bg-light">
					<div className="row my-5">
						<div className="col text-center">
							<h1 className="display-1">pptxgenjs</h1>
						</div>
					</div>
					<div className="row my-3">
						<div className="col text-center">
							<h4>Fastest growing powerpoint library</h4>
						</div>
					</div>
				</header>
				<main>
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
