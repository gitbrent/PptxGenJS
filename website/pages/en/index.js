/**
 * Copyright (c) 2017-present, Facebook, Inc.
 *
 * This source code is licensed under the MIT license found in the
 * LICENSE file in the root directory of this source tree.
 */

const React = require('react');
const CompLibrary = require('../../core/CompLibrary.js');
const MarkdownBlock = CompLibrary.MarkdownBlock; /* Used to read markdown */
const Container = CompLibrary.Container;
const GridBlock = CompLibrary.GridBlock;
const siteConfig = require(process.cwd() + '/siteConfig.js');

function imgUrl(img) {
	return siteConfig.baseUrl + 'img/' + img;
}

function docUrl(doc, language) {
	return siteConfig.baseUrl + 'docs/' + (language ? language + '/' : '') + doc;
}

function pageUrl(page, language) {
	return siteConfig.baseUrl + (language ? language + '/' : '') + page;
}

class Button extends React.Component {
	render() {
		return (
			<div className="pluginWrapper buttonWrapper">
				<a className="button" href={this.props.href} target={this.props.target}>
					{this.props.children}
				</a>
			</div>
		);
	}
}

Button.defaultProps = {
	target: '_self',
};

const SplashContainer = props => (
	<div className="homeContainer">
		<div className="homeSplashFade">
			<div className="wrapper homeWrapper">{props.children}</div>
		</div>
	</div>
);

const Logo = props => (
	<div className="projectLogo">
		<img src={props.img_src} />
	</div>
);

const ProjectTitle = props => (
	<h2 className="projectTitle">
		{siteConfig.title}
		<small>{siteConfig.tagline}</small>
	</h2>
);

const PromoSection = props => (
	<div className="section promoSection">
		<div className="promoRow">
			<div className="pluginRowBlock">{props.children}</div>
		</div>
	</div>
);

/* ========== */

function MakeLeftBulletText(strText) {
	return '<p style="text-align:left">&bull; '+ strText +'</p>';
}

/* ========== */

class HomeSplash extends React.Component {
	render() {
		let language = this.props.language || '';
		return (
			<SplashContainer>
				<Logo img_src={imgUrl('pptxgenjs.svg')} />
				<div className="inner">
					<ProjectTitle />
					<PromoSection>
						<Button href="#try">Try It Out</Button>
						<Button href={docUrl('doc1.html', language)}>Get Started</Button>
					</PromoSection>
				</div>
			</SplashContainer>
		);
	}
}

const Block = props => (
	<Container
		padding={['bottom', 'top']}
		id={props.id}
		background={props.background}>
		<GridBlock align="center" contents={props.children} layout={props.layout} />
	</Container>
);

const Features = props => (
	<Block background="light" layout="fourColumn">
		{[
			{
				content: 'Works with all current web browsers (Chrome, Edge, Firefox, etc.) and IE11',
				image: imgUrl('circle-handshake.svg'),
				imageAlign: 'top',
				title: 'Widely Supported',
			},
			{
				content: 'Charts, Images, Media, Shapes, Tables and Text (plus Master Slides)',
				image: imgUrl('circle-art.svg'),
				imageAlign: 'top',
				title: 'Full Featured',
			},
			{
				content: 'Entire PowerPoint presentations can be created in a few lines of code',
				image: imgUrl('circle-magic.svg'),
				imageAlign: 'top',
				title: 'Easy To Use',
			},
			{
				content: 'Pure JavaScript solution - everything necessary to create PowerPoint PPT exports is included',
				image: imgUrl('circle-blueprint.svg'),
				imageAlign: 'top',
				title: 'Modern',
			},
		]}
	</Block>
);

const FeatureCallout = props => (
	<Block background="white">
		{[
			{
				image: imgUrl('pptxgenjs.svg'),
				imageAlign: 'right',
				title: 'Additional Features',
				content: '<ul style="text-align:left;">'
					+ '<li>'+'Support for all major PowerPoint shapes, including charts'+'</li>'
					+ '<li>'+'Define any size Slide (e.g.: A4)'+'</li>'
					+ '<li>'+'Supports corporate/brand Slide Master designs'+'</li>'
					+ '<li>'+'Supports RTL (right-to-left) text'+'</li>'
					+ '<li>'+'Supports Chinese and other international language/fonts'+'</li>'
					+ '<li>'+'Node.js can utilize callbacks, binary streaming, and more'+'</li>'
					+ '<li>'+'Works with Electron applications'+'</li>'
					+ '</ul>'
			},
		]}
	</Block>
);

const TryOut = props => (
	<Block background="light" id="try">
		{[
			{
				title: 'Try it Out',
				image: imgUrl('demo-simple.png'),
				imageAlign: 'left',
				content: 'Live demo',
			},
		]}
	</Block>
);

const LearnMore = props => (
	<Block background="white">
		{[
			{
				title: 'Learn More',
				image: imgUrl('docusaurus.svg'),
				imageAlign: 'right',
				content: 'Here is a ton of docs and examples!',
			},
		]}
	</Block>
);


class Index extends React.Component {
	render() {
		let language = this.props.language || '';

		return (
			<div>
				<HomeSplash language={language} />
				<div className="mainContainer">
					<Features />
					<FeatureCallout />
					<TryOut />
					<LearnMore />
				</div>
			</div>
		);
	}
}

module.exports = Index;
