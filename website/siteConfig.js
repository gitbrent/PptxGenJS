/**
 * Copyright (c) 2017-present, Brent Ely
 *
 * This source code is licensed under the MIT license found in the
 * LICENSE file in the root directory of this source tree.
 */

const siteConfig = {
	title: 'PptxGenJS',
	tagline: 'JavaScript library that creates PowerPoint presentations',
	url: 'https://gitbrent.github.io',
	baseUrl: '/PptxGenJS/',
	projectName: 'PptxGenJS',
	gaTrackingId: 'UA-75147115-1',
	headerLinks: [
		{href: 'https://gitbrent.github.io/PptxGenJS/releases', label: 'Download'},
		{doc: 'installation', label: 'Get Started'},
		{doc: 'usage-basic-create', label: 'API'},
		{page: 'help', label: 'Help'},
		{href: 'https://gitbrent.github.io/PptxGenJS/', label: 'GitHub'},
	],
	headerIcon: 'img/pptxgenjs.svg',
	footerIcon: 'img/pptxgenjs.svg',
	favicon: 'img/favicon.png',
	colors: {
		primaryColor: '#de4b2c',
		secondaryColor: '#bf360c',
	},
	copyright: 'Copyright © '+ new Date().getFullYear() +' Brent Ely',
	projectName: 'PptxGenJS',
	highlight: {
		theme: 'hybrid',
		defaultLang: 'javascript',
	},
	scripts: [
		'https://cdn.rawgit.com/gitbrent/PptxGenJS/v2.0.0/dist/pptxgen.bundle.js',
		'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/9.12.0/highlight.min.js',
	],
	repoUrl: 'https://github.com/gitbrent/PptxGenJS',
};

module.exports = siteConfig;
