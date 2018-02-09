/**
 * Copyright (c) 2017-present, Brent Ely
 *
 * This source code is licensed under the MIT license found in the
 * LICENSE file in the root directory of this source tree.
 */

/*
NOTE: FIXME:

`onClick` binding does not work. I used sample code from react.com and Docusaurus strips all `onclick` events apparretny;
<button onClick={(e) => this.handleClick(e)}>YOYOYO</button>
*/

const React = require('react');

class Footer extends React.Component {
	/*
	constructor(props) {
		super(props);
		this.handleClick = this.handleClick.bind(this);
	}
	*/

	docUrl(doc, language) {
		const baseUrl = this.props.config.baseUrl;
		return baseUrl + 'docs/' + (language ? language + '/' : '') + doc;
	}

	pageUrl(doc, language) {
		const baseUrl = this.props.config.baseUrl;
		return baseUrl + (language ? language + '/' : '') + doc;
	}

	/*
	handleClick() {
		console.log('The link was clicked.');
	}
	*/

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
						<a
							href={this.docUrl('installation.html', this.props.language)}
							onClick={()=>ga('send','event','Link','click','Link-installation')}>
							Getting Started With PptxGenJS
						</a>
						<a
							href={this.docUrl('usage-basic-create.html', this.props.language)}
							onClick={this.handleClick}>
							PowerPoint Library API Reference
						</a>
						<a href={this.docUrl('usage-basic-create.html', this.props.language)}>
							PowerPoint JavaScript Code Samples
						</a>
					</div>
					<div>
						<h5>Community</h5>
						<a
							href="https://github.com/gitbrent/pptxgenjs/issues"
							target="_blank">
							GitHub Issues
						</a>
						<a
							href="http://stackoverflow.com/questions/tagged/pptxgenjs"
							target="_blank">
							Stack Overflow
						</a>
					</div>
					<div>
						<h5>More</h5>
						<a href="https://github.com/gitbrent/pptxgenjs">GitHub Project</a>
						<a href={this.props.config.baseUrl + 'blog'}>Blog</a>
						<a href="https://www.flaticon.com/packs/creativity">Site Icons</a>
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
