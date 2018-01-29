/**
 * Copyright (c) 2017-present, Facebook, Inc.
 * Copyright (c) 2017-present, Brent Ely
 *
 * This source code is licensed under the MIT license found in the
 * LICENSE file in the root directory of this source tree.
 */

const React = require('react');

class Footer extends React.Component {
	docUrl(doc, language) {
		const baseUrl = this.props.config.baseUrl;
		return baseUrl + 'docs/' + (language ? language + '/' : '') + doc;
	}

	pageUrl(doc, language) {
		const baseUrl = this.props.config.baseUrl;
		return baseUrl + (language ? language + '/' : '') + doc;
	}

	render() {
		const currentYear = new Date().getFullYear();
		return (
			<footer className="nav-footer" id="footer">
				<section className="sitemap">
					<a href={this.props.config.baseUrl} className="nav-home">
						{this.props.config.footerIcon && (
							<img
								src={this.props.config.baseUrl + this.props.config.footerIcon}
								alt={this.props.config.title}
								width="66"
								height="58"
							/>
						)}
					</a>
					<div>
						<h5>Docs</h5>
						<a href={this.docUrl('doc1.html', this.props.language)}>
							Getting Started (or other categories)
						</a>
						<a href={this.docUrl('doc2.html', this.props.language)}>
							Guides (or other categories)
						</a>
						<a href={this.docUrl('doc3.html', this.props.language)}>
							API Reference (or other categories)
						</a>
					</div>
					<div>
						<h5>Community</h5>
						<a href={this.pageUrl('users.html', this.props.language)}>
							User Showcase
						</a>
						<a
							href="http://stackoverflow.com/questions/tagged/pptxgenjs"
							target="_blank">
							Stack Overflow
						</a>

					</div>
					<div>
						<h5>More</h5>
						<a href="https://www.flaticon.com/packs/creativity">SVG Icons</a>
						<a href={this.props.config.baseUrl + 'blog'}>Blog</a>
						<a href="https://github.com/gitbrent/pptxgenjs">GitHub</a>
						<a
							className="github-button"
							href={this.props.config.repoUrl}
							data-icon="octicon-star"
							data-count-href="/gitbrent/pptxgenjs/stargazers"
							data-show-count={true}
							data-count-aria-label="# stargazers on GitHub"
							aria-label="Star this project on GitHub">
							Star
						</a>
					</div>
				</section>

				<section className="copyright">
					Copyright &copy; {currentYear} Brent Ely
				</section>
			</footer>
		);
	}
}

module.exports = Footer;
