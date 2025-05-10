// Inside your plugins/pptxgenjs-plugin/index.js
const webpack = require('webpack');

export default function(context, options) {
	return {
		name: 'pptxgenjs-plugin',
		configureWebpack: (config, isServer, utils) => {
			const {
				// Docusaurus Webpack utilities can be accessed here if needed
			} = utils;

			return {
				plugins: [
					// This plugin will replace the 'node:' prefix with an empty string
					// for the specified modules.
					new webpack.NormalModuleReplacementPlugin(/^node:(fs|https|path|os|image-size|fs\/promises)$/, (resource) => {
						resource.request = resource.request.replace(/^node:/, '');
					}),
				],
				resolve: {
					fallback: {
						...config.resolve.fallback,
						// Add 'fs/promises' to the fallback
						"fs": false,
						"https": false,
						"image-size": false,
						"os": false,
						"path": false,
						"fs/promises": false, // Explicitly ignore fs/promises
						// No need for node:fs, node:https etc. here if using NormalModuleReplacementPlugin
					},
				},
			};
		}
	};
}
