/**
 * Copyright (c) 2017-present, Facebook, Inc.
 *
 * This source code is licensed under the MIT license found in the
 * LICENSE file in the root directory of this source tree.
 */

/* List of projects/orgs using your project for the users page */
const users = [
	{
		caption: 'PptxGenJS',
		image: '/PptxGenJS/img/pptxgenjs.svg',
		infoLink: 'https://gitbrent.github.io/PptxGenJS/',
		pinned: true,
	},
];

const siteConfig = {
	title: 'PptxGenJS' /* title for your website */,
	tagline: 'JavaScript library that creates PowerPoint presentations',
	url: 'https://gitbrent.github.io' /* your website url */,
	baseUrl: '/PptxGenJS/' /* base url for your project */,
	projectName: 'PptxGenJS',
	headerLinks: [
		{href: 'https://gitbrent.github.io/PptxGenJS/releases', label: 'Download'},
		{doc: 'doc1', label: 'Get Started'},
		{doc: 'doc4', label: 'API'},
		{page: 'help', label: 'Help'},
		{href: 'https://gitbrent.github.io/PptxGenJS/', label: 'GitHub'},
	],
	users,
	/* path to images for header/footer */
	headerIcon: 'img/pptxgenjs.svg',
	footerIcon: 'img/pptxgenjs.svg',
	favicon: 'img/favicon.png',
	/* colors for website */
	colors: {
		primaryColor: '#DE4B2C',
		secondaryColor: '#205C3B',
	},
	// This copyright info is used in /core/Footer.js and blog rss/atom feeds.
	copyright:
		'Copyright © ' +
		new Date().getFullYear() +
		' Brent Ely',
	// organizationName: 'deltice', // or set an env variable ORGANIZATION_NAME
	projectName: 'PptxGenJS', // or set an env variable PROJECT_NAME
	highlight: {
		// Highlight.js theme to use for syntax highlighting in code blocks
		theme: 'atom-one-dark',
		defaultLang: 'javascript',
	},
	scripts: ['https://buttons.github.io/buttons.js'],
	repoUrl: 'https://github.com/gitbrent/PptxGenJS',
	gaTrackingId: 'UA-75147115-1',
};

module.exports = siteConfig;
