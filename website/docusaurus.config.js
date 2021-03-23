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
	scripts: ["https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@latest/dist/pptxgen.bundle.js", "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/9.12.0/highlight.min.js"],
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
		image: "img/pptxgenjs.png",
		navbar: {
			title: "PptxGenJS",
			logo: {
				alt: "PptGenJS Logo",
				src: "img/pptxgenjs.svg",
				srcDark: "img/pptxgenjs.svg",
			},
			items: [
				{
					to: "docs/",
					label: "Get Started",
					position: "left",
				},
				{
					to: "docs/installation",
					label: "Documentation",
					position: "left",
				},
				{
					href: "https://github.com/gitbrent/PptxGenJS/releases",
					label: "Download",
					position: "left",
				},
				{
					href: "https://github.com/gitbrent/PptxGenJS",
					position: "right",
					className: "header-github-link",
					"aria-label": "GitHub repository",
				},
				{
					href: "https://github.com/gitbrent/PptxGenJS/",
					label: "GitHub",
					position: "right",
				},
			],
		},
		footer: {
			style: "dark",
			links: [
				{
					title: "Learn",
					items: [
						{
							label: "Introduction",
							to: "docs/quick-start",
						},
						{
							label: "Installation",
							to: "docs/installation",
						},
						{
							label: "JSFiddle Demo",
							href: "https://jsfiddle.net/gitbrent/L1uctxm0/",
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
				src: "img/pptxgenjs.svg",
				href: "https://github.com/gitbrent/PptxGenJS",
			},
		},
		gtag: {
			trackingID: "UA-75147115-1",
		},
	},
};
