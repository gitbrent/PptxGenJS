import typescript from 'rollup-plugin-typescript2'
import pkg from './package.json'

export default {
	input: 'src/pptxgen.ts',
	output: [
		{
			file: './src/bld/pptxgen.js',
			format: 'iife',
			name: 'PptxGenJS',
			globals: {
				jszip: 'JSZip'
			}
		},
		{
			file: './src/bld/pptxgen.cjs.js',
			format: 'cjs'
		},
		/*
		{
			file: './src/bld/pptxgenjs.umd.js',
			format: 'umd',
			name: 'PptxGenJS',
			globals: {
				jszip: 'JSZip'
			}
		},
		*/
		{
			file: './src/bld/pptxgen.es.js',
			format: 'es'
		}
	],
	external: [...Object.keys(pkg.dependencies || {}), ...Object.keys(pkg.peerDependencies || {})],
	plugins: [
		typescript({
			typescript: require('typescript')
		})
	]
}
