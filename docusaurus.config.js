const lightCodeTheme = require("prism-react-renderer/themes/github");
const darkCodeTheme = require("prism-react-renderer/themes/dracula");

// With JSDoc @type annotations, IDEs can provide config autocompletion
/** @type {import('@docusaurus/types').DocusaurusConfig} */
(
	module.exports = {
		title: "PptxGenJS",
		tagline: "Create JavaScript PowerPoint Presentations",
		url: "https://gitbrent.github.io",
		baseUrl: "/PptxGenJS/",
		organizationName: "PptxGenJS",
		projectName: "PptxGenJS",
		baseUrlIssueBanner: true,
		url: "https://gitbrent.github.io",
		customFields: {
			repoUrl: "https://github.com/gitbrent/PptxGenJS",
		},
		onBrokenLinks: "throw",
		onBrokenMarkdownLinks: "warn",
		trailingSlash: true,
		favicon: "img/favicon.png",
		presets: [
			[
				"@docusaurus/preset-classic",
				/** @type {import('@docusaurus/preset-classic').Options} */
				({
					docs: {
						showLastUpdateAuthor: true,
						showLastUpdateTime: true,
						path: "./docs",
						sidebarPath: require.resolve("./sidebars.js"),
					},
					/*
					blog: {
						showReadingTime: true,
						// Please change this to your repo.
						editUrl: "https://github.com/facebook/docusaurus/edit/main/website/blog/",
					},
					*/
					theme: {
						customCss: require.resolve("./src/css/custom.css"),
					},
				}),
			],
		],
		themeConfig: {
			liveCodeBlock: {
				playgroundPosition: "bottom",
			},
			hideableSidebar: true,
			colorMode: {
				defaultMode: "light",
				disableSwitch: false,
				respectPrefersColorScheme: true,
			},
			announcementBar: {
				id: "supportus",
				content: '⭐️  If you like PptxGenJS, give it a star on <a target="_blank" rel="noopener noreferrer" href="https://github.com/gitbrent/PptxGenJS">GitHub</a>! ⭐️',
			},
			prism: {
				theme: lightCodeTheme,
				darkTheme: darkCodeTheme,
			},
			image: "img/app-gears.svg",
			navbar: {
				style: "dark",
				title: "PptxGenJS",
				logo: {
					alt: "PptGenJS Logo",
					src: "img/app-gears.svg",
					srcDark: "img/app-gears.svg",
				},
				items: [
					{
						to: "docs/quick-start",
						label: "Get Started",
						position: "left",
						"aria-label": "get started",
					},
					/*{
						to: "docs/installation",
						label: "Installation",
						position: "left",
						"aria-label": "installation",
					},*/
					{
						to: "demos",
						label: "Demos",
						position: "left",
						"aria-label": "demos",
					},
					{
						to: "html2pptx",
						label: "HTML-to-PPTX",
						position: "left",
						"aria-label": "html-to-pptx",
					},
					{
						href: "https://github.com/gitbrent/PptxGenJS/releases",
						label: "Latest Release",
						position: "left",
						"aria-label": "latest release",
					},
					{
						to: "sponsor",
						label: "Donate",
						position: "left",
						className: "navbar-sponsor icon-svg-coin",
						"aria-label": "sponsor us with a donation",
					},
					/*{
						href: "https://www.npmjs.com/package/pptxgenjs",
						position: "right",
						className: "header-npm-link",
						"aria-label": "NPM homepage",
					},
					{
						href: "https://github.com/gitbrent/PptxGenJS",
						position: "right",
						className: "header-github-link",
						"aria-label": "GitHub repository",
					},*/
					{
						href: "https://www.npmjs.com/package/pptxgenjs",
						label: "npm",
						position: "right",
						"aria-label": "npm home page",
					},
					{
						href: "https://github.com/gitbrent/PptxGenJS/",
						label: "GitHub",
						position: "right",
						"aria-label": "GitHub repository",
					},
				],
			},
			footer: {
				style: "light",
				links: [
					{
						title: "Learn",
						items: [
							{
								label: "Quick Start",
								to: "docs/quick-start",
							},
							{
								label: "Installation",
								to: "docs/installation",
							},
							{
								label: "Demos",
								href: "/demos",
							},
						],
					},
					{
						title: "Community",
						items: [
							{
								label: "Stack Overflow",
								href: "https://stackoverflow.com/questions/tagged/pptxgenjs",
							},
						],
					},
					{
						title: "More",
						items: [
							{
								label: "GitHub",
								href: "https://github.com/gitbrent/pptxgenjs",
							},
							{
								label: "Twitter",
								href: "https://twitter.com/pptxgenjs",
							},
						],
					},
					{
						title: "Legal",
						items: [
							{
								label: "Privacy",
								href: "/privacy",
							},
							{
								label: "License",
								href: "/license",
							},
						],
					},
				],
				copyright: `Copyright © ${new Date().getFullYear()} Brent Ely`,
				logo: {
					alt: "PptxGenJS Logo",
					src: "img/pptxgenjs-footer.png",
					href: "https://github.com/gitbrent/PptxGenJS",
				},
			},
			gtag: {
				trackingID: "G-4F7ZC3PH3Y",
			},
		},
	}
);
