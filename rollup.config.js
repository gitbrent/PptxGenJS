import typescript from 'rollup-plugin-typescript2'
import { terser } from 'rollup-plugin-terser'
import banner from 'rollup-plugin-banner'
import clear from 'rollup-plugin-clear'
import pkg from './package.json'

const BANNER = `bulletpoints
v<%= pkg.version %> (${new Date().toISOString()}) 
2019-${new Date().getFullYear()} <%= pkg.author %>
`

export default {
    input: 'src/pptxgen.ts',
    output: [
        {
            file: './dist/bulletpoints.bundle.js',
            format: 'iife',
            name: 'PptxGenJS',
            globals: {
                jszip: 'JSZip',
                'calc-units': 'calcUnits'
            }
        },
        {
            file: pkg.main,
            format: 'cjs'
        },
        {
            file: pkg.module,
            format: 'es'
        }
    ],
    external: [
        ...Object.keys(pkg.dependencies || {}),
        ...Object.keys(pkg.peerDependencies || {})
    ],

    plugins: [
        clear({
            // required, point out which directories should be clear.
            targets: ['dist', 'types']
        }),
        typescript({
            typescript: require('typescript'),
            useTsconfigDeclarationDir: true
        }),
        terser({
            include: [/^.+\.bundle\.js$/]
        }),
        banner(BANNER)
    ]
}
