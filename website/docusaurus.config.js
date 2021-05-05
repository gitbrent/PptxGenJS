/** @type {import('@docusaurus/types').DocusaurusConfig} */
module.exports = {
	title: "PptxGenJS",
	tagline: "Create JavaScript PowerPoint Presentations",
	url: "https://gitbrent.github.io",
	baseUrl: "/PptxGenJS/",
	organizationName: "PptxGenJS",
	projectName: "PptxGenJS",
	baseUrlIssueBanner: true,
	url: "https://gitbrent.github.io",
	onBrokenLinks: "throw",
	onBrokenMarkdownLinks: "warn",
	favicon: "img/favicon.png",
	customFields: {
		repoUrl: "https://github.com/gitbrent/PptxGenJS",
	},
	onBrokenLinks: "log",
	onBrokenMarkdownLinks: "log",
	presets: [
		[
			"@docusaurus/preset-classic",
			{
				// Debug defaults to true in dev, false in prod
				debug: undefined,
				// Will be passed to @docusaurus/theme-classic.
				theme: {
					customCss: [require.resolve("./src/css/custom.css")],
				},
				docs: {
					showLastUpdateAuthor: true,
					showLastUpdateTime: true,
					path: "./docs",
					sidebarPath: "./sidebars.json",
				},
				blog: false,
			},
		],
	],
	plugins: [
		[
			"@docusaurus/plugin-client-redirects",
			{
				fromExtensions: ["html"],
			},
		],
	],
	themes: ["@docusaurus/theme-live-codeblock"],
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
			theme: require("prism-react-renderer/themes/github"),
			darkTheme: require("prism-react-renderer/themes/dracula"),
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
				},
				{
					to: "docs/installation",
					label: "Installation",
					position: "left",
				},
				{
					to: "pptxdemos",
					label: "Demos",
					position: "left",
				},
				{
					href: "https://github.com/gitbrent/PptxGenJS/releases",
					label: "Latest Release",
					position: "left",
				},
				{
					to: "sponsor",
					label: "Sponsor Us",
					position: "left",
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
							href: "pptxdemos",
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
};
