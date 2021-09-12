import React from "react";
import Layout from "@theme/Layout";
//import "../css/bootstrap-yeti.css";
import "../css/purged.css";

const ADDRESS_BTC = "bc1qm8cqunm00wtxe7eztspfkaysc6n7fedrsvfv6c";
const ADDRESS_ETH = "0x8F130f0522c9E185E0692B0B2801295732951bAA";
const ADDRESS_DOT = "112BhnEjhsVPEDPtLqx2pKjuZmu64GDxFHLaxdbqwpu5h8ME";
const ADDRESS_CRO = "cro1zlgepgq383k43auqha5lp35j8y0jdtku2dm3t6";
const ADDRESS_DOGE = "DDJjv6fsfCwxsyyHPSZrSJMZC51LrUtav7";

export default () => {
	const PageHeader = () => {
		return (
			<section className="useTheme mb-5">
				<div className="card border-0">
					<div className="card-header bg-success text-white">
						<h5 className="mb-0">Support this project by making a donation!</h5>
					</div>
					<div className="card-body">
						<p className="card-text">Are you building killer apps? Showcasing your amazing research? Selling products that include our library?</p>
						<p className="card-text">Consider making a contribution to sponsor library development.</p>
						<p className="card-text">❤️&nbsp; Thanks to all of our sponsors and contributors! ❤️</p>
					</div>
				</div>
			</section>
		);
	};

	const CryptoCards = () => {
		// NOTE: 20210407: cant use `row-cols-12` yet as docusaurus core bootstrap is ruining `rows` style, use "col-12" etc on cols to supercede for now
		return (
			<section className="useTheme">
				<div className="row g-5 mb-5">
					<div className="col-12 col-md-3">
						<div className="bgTheme text-center h-100 p-4">
							<h4 className="mb-4">Bitcoin (BTC)</h4>
							<p>
								<img src="/PptxGenJS/img/sponsor_btc.png" alt="bitcoin wallet address" />
							</p>
							<div className="row justify-content-center align-items-center g-3">
								<div className="col-auto font-monospace">{ADDRESS_BTC}</div>
								<div className="col-auto">
									<button
										type="button"
										title="click to copy"
										class="btn btn-sm btn-primary"
										onClick={() => {
											navigator.clipboard.writeText(ADDRESS_BTC);
										}}
									>
										Copy
									</button>
								</div>
							</div>
						</div>
					</div>
					<div className="col-12 col-md-3">
						<div className="bgTheme text-center h-100 p-4">
							<h4 className="mb-4">Etherium (ETH)</h4>
							<p>
								<img src="/PptxGenJS/img/sponsor_eth.png" alt="etherium wallet address" />
							</p>
							<div className="row justify-content-center align-items-center g-3">
								<div className="col-auto font-monospace">{ADDRESS_ETH}</div>
								<div className="col-auto">
									<button
										type="button"
										title="click to copy"
										class="btn btn-sm btn-primary"
										onClick={() => {
											navigator.clipboard.writeText(ADDRESS_ETH);
										}}
									>
										Copy
									</button>
								</div>
							</div>
						</div>
					</div>
					<div className="col-12 col-md-3">
						<div className="bgTheme text-center h-100 p-4">
							<h4 className="mb-4">Doge (DOGE)</h4>
							<p>
								<img src="/PptxGenJS/img/sponsor_doge.png" alt="doge wallet address" />
							</p>
							<div className="row justify-content-center align-items-center g-3">
								<div className="col-auto font-monospace">{ADDRESS_DOGE}</div>
								<div className="col-auto">
									<button
										type="button"
										title="click to copy"
										class="btn btn-sm btn-primary"
										onClick={() => {
											navigator.clipboard.writeText(ADDRESS_DOGE);
										}}
									>
										Copy
									</button>
								</div>
							</div>
						</div>
					</div>
					<div className="col-12 col-md-3">
						<div className="bgTheme text-center h-100 p-4">
							<h4 className="mb-4">Crypto.com (CRO)</h4>
							<p>
								<img src="/PptxGenJS/img/sponsor_cro.png" alt="crypto.com wallet address" />
							</p>
							<div className="row justify-content-center align-items-center g-3">
								<div className="col-auto font-monospace">{ADDRESS_CRO}</div>
								<div className="col-auto">
									<button
										type="button"
										title="click to copy"
										class="btn btn-sm btn-primary"
										onClick={() => {
											navigator.clipboard.writeText(ADDRESS_CRO);
										}}
									>
										Copy
									</button>
								</div>
							</div>
						</div>
					</div>
				</div>
			</section>
		);
	};

	return (
		<Layout title="Sponsor Us">
			<div className="container my-4">
				<h1 className="mb-4">Donate / Sponsor</h1>
				<PageHeader />
				<CryptoCards />
			</div>
		</Layout>
	);
};
