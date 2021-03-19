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
						<h5>Documentation</h5>
						<a
							href={this.docUrl('quick-start.html')}
							onClick={()=>ga('send','event','Link','click','link-footer-GetStarted')}>
							Getting Started With PptxGenJS
						</a>
						<a
							href={this.docUrl('installation.html')}
							onClick={()=>ga('send','event','Link','click','link-footer-LibraryApi')}>
							PowerPoint Library API Reference
						</a>
						<a href={this.docUrl('usage-pres-create.html')}
							onClick={()=>ga('send','event','Link','click','link-footer-CodeSamples')}>
							PowerPoint JavaScript Code Samples
						</a>
					</div>
					<div>
						<h5>More</h5>
						<a
							href="https://jsfiddle.net/gitbrent/L1uctxm0/"
							target="_blank">
							JSFiddle Demo Presentation
						</a>
						<a
							href="https://github.com/gitbrent/pptxgenjs/issues"
							target="_blank">
							PptxGenJS GitHub Issues
						</a>
						<a
							href="http://stackoverflow.com/questions/tagged/pptxgenjs"
							target="_blank">
							PptxGenJS on Stack Overflow
						</a>
					</div>
					<div>
						<h5>Social</h5>
						<a href="https://twitter.com/pptxgenjs">Twitter</a>
						<a href="https://www.pinterest.com/pptxgenjs">Pinterest</a>
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
