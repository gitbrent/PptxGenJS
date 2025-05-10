/**
 * https://docusaurus.io/docs/sidebar/items
 */

export default {
	docs: [
		{
			type: 'doc',
			id: 'introduction',
			label: 'Introduction',
		},
		{
			type: 'category',
			label: 'Get Started',
			collapsible: true,
			collapsed: false,
			items: ['quick-start', 'installation', 'integration'],
		},
		{
			type: 'category',
			label: 'Usage',
			collapsible: true,
			collapsed: true,
			items: ['usage-pres-create', 'usage-pres-options', 'usage-add-slide', 'usage-slide-options', 'usage-saving'],
		},
		{
			type: 'category',
			label: 'Features',
			collapsible: true,
			collapsed: true,
			items: ['html-to-powerpoint', 'masters', 'sections', 'shapes-and-schemes', 'speaker-notes', 'types'],
		},
		{
			type: 'category',
			label: 'API Reference',
			collapsible: true,
			collapsed: true,
			items: ['api-charts', 'api-images', 'api-media', 'api-shapes', 'api-tables', 'api-text'],
		},
		{
			type: 'category',
			label: 'Misc',
			collapsible: true,
			collapsed: true,
			items: ['deprecated'],
		},
	],
};
