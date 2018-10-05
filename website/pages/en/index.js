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

/* ========== */

function MakeLeftBulletText(strText) {
	return '<p style="text-align:left">&bull; '+ strText +'</p>';
}

// NOTE: Code is only recognized if lines have leading tabs (?)
const tryCodeBlock = `
	var pptx = new PptxGenJS();
	var slide = pptx.addNewSlide();
	slide.addText(
	    "BONJOUR - CIAO - GUTEN TAG - HELLO - HOLA - NAMASTE - OLÀ - ZDRAS-TVUY-TE - こんにちは - 你好",
	    { x:0, y:1, w:'100%', h:2, align:'c', color:'0088CC', fill:'F1F1F1', fontSize:24 }
	);
	pptx.save('PptxGenJS-Demo');
`;

/* ========== */

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

function imgUrl(img) {
	return siteConfig.baseUrl + 'img/' + img;
}

function docUrl(doc, language) {
	return siteConfig.baseUrl + 'docs/' + (language ? language + '/' : '') + doc;
}

function pageUrl(page, language) {
	return siteConfig.baseUrl + (language ? language + '/' : '') + page;
}

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

const Block = props => (
	<Container
		padding={['bottom', 'top']}
		id={props.id}
		background={props.background}>
		<GridBlock align={props.align||"center"} contents={props.children} layout={props.layout} />
	</Container>
);

/* ============================== */

// 1: Top
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
						<Button href={docUrl('installation.html', language)}>Get Started</Button>
					</PromoSection>
				</div>
			</SplashContainer>
		);
	}
}

// 2:
const FeatureBullets = props => (
	<Block background="light" layout="fourColumn">
		{[
			{
				content: 'Works with all current web browsers (Chrome, Edge, Firefox, etc.) and IE11',
				image: imgUrl('circle-handshake.svg'),
				imageAlign: 'top',
				title: 'Widely Supported',
			},
			{
				content: 'Create charts, images, media, shapes, tables, text and utilize Master Slides',
				image: imgUrl('circle-art.svg'),
				imageAlign: 'top',
				title: 'Full Featured',
			},
			{
				content: 'Entire PowerPoint presentations can be created using just a few lines of code',
				image: imgUrl('circle-magic.svg'),
				imageAlign: 'top',
				title: 'Easy To Use',
			},
			{
				content: 'Pure JavaScript solution that works with browsers, Node, Angular, Electron and more',
				image: imgUrl('circle-blueprint.svg'),
				imageAlign: 'top',
				title: 'Modern',
			},
		]}
	</Block>
);

// 3:
const FeatureCallout = props => (
	<Block background="white" id="FeatureBullets">
		{[
			{
				image: imgUrl('feature-callout.png'),
				imageAlign: 'right',
				title: 'Additional Features',
				content: '<ul style="text-align:left;">'
					+ '<li>'+'Support for all major PowerPoint chart types'+'</li>'
					+ '<li>'+'Support for custom Slide sizes (A4, etc.)'+'</li>'
					+ '<li>'+'Support for corporate/brand Slide Master designs'+'</li>'
					+ '<li>'+'Support for RTL (right-to-left) text'+'</li>'
					+ '<li>'+'Support for Chinese and other international language/fonts'+'</li>'
					+ '<li>'+'Compatible with Node, Angular, Electron and other application frameworks'+'</li>'
					+ '<li>'+'Node and other libraries/apps can use advanced export types such as callbacks and binary streaming'+'</li>'
					+ '</ul>'
			},
		]}
	</Block>
);

// 4:
const TryOutLiveDemo = props => (
	<Block background="light" id="try" align="left" layout="oneColumn">
		{[
			{
				title: 'Live Demo: Create a simple pptx presentation',
				content: tryCodeBlock,
			},
			{
				title: '',
				content: '<p>Any desktop or mobile browser that is capable of downloading files can execute the code above to create a presentation.</p>'
				+ `<Button class="button" href="javascript:" onclick="eval(document.getElementById('try').getElementsByClassName('hljs')[0].innerText); if(ga)ga('send','event','Link','click','Demo-Simple');">Try It Out</Button>`
				+ '<br/><br/><p>There is also a pre-configured <a href="https://jsfiddle.net/gitbrent/gx34jy59/5/">jsFiddle demo</a> available.</p>',
			},
		]}
	</Block>
);

// 5:
// TODO: <li>View sample code and PowerPoint presentations</li>
const LearnMore = props => (
	<Block background="white" id="learn">
		{[
			{
				title: 'Learn More',
				image: imgUrl('learn-more.png'),
				imageAlign: 'left',
				content: '<ul style="text-align:left">'
					+ '<li><a href="'+ docUrl('installation.html', '') +'" '
					+ ' onclick="if(ga)ga(\'send\',\'event\',\'Link\',\'click\',\'LearnMore-installation\')">'
					+ 'Installing PptxGenJS</a></li>'
					+ '<li><a href="'+ docUrl('installation.html', '') +'" '
					+ ' onclick="if(ga)ga(\'send\',\'event\',\'Link\',\'click\',\'LearnMore-installation\')">'
					+ 'Creating a Presentation</a></li>'
					+ '<li><a href="'+ docUrl('masters.html', '') +'" '
					+ ' onclick="if(ga)ga(\'send\',\'event\',\'Link\',\'click\',\'LearnMore-masters\')">'
					+ 'Master Slides and Layouts</a></li>'
					+ '<li><a href="'+ docUrl('api-tables.html', '') +'" '
					+ ' onclick="if(ga)ga(\'send\',\'event\',\'Link\',\'click\',\'LearnMore-api-tables\')">'
					+ 'View Table API and sample code</a></li>'
					+ '<li><a href="'+ docUrl('html-to-powerpoint.html', '') +'" '
					+ ' onclick="if(ga)ga(\'send\',\'event\',\'Link\',\'click\',\'LearnMore-table2slides\')">'
					+ 'Converting HTML tables to presentations</a></li>'
					+ '</ul>'
			},
		]}
	</Block>
);

// DEFINE PAGE
class Index extends React.Component {
	render() {
		let language = this.props.language || '';

		return (
			<div>
				<HomeSplash language={language} />
				<div className="mainContainer">
					<FeatureBullets />
					<FeatureCallout />
					<TryOutLiveDemo />
					<LearnMore />
				</div>
				<script>hljs.initHighlightingOnLoad();</script>
			</div>
		);
	}
}

module.exports = Index;
