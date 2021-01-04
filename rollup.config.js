import pkg from "./package.json";
import resolve from "@rollup/plugin-node-resolve";
import commonjs from "@rollup/plugin-commonjs";
import typescript from "rollup-plugin-typescript2";

export default {
  input: "src/pptxgen.ts",
  output: [
    {
      file: "./src/bld/pptxgen.js",
      format: "iife",
      name: "PptxGenJS",
      globals: {
        jszip: "JSZip",
      },
    },
    {
      file: "./src/bld/pptxgen.cjs.js",
      format: "cjs",
      exports: "default",
    },
    /*
    {
      file: "./src/bld/pptxgen.umd.js",
      format: "umd",
      name: "PptxGenJS",
      globals: {
        jszip: "JSZip",
      },
	},
	*/
    {
      file: "./src/bld/pptxgen.es.js",
      format: "es",
    },
  ],
  external: [...Object.keys(pkg.dependencies || {}), ...Object.keys(pkg.peerDependencies || {})],
  plugins: [
    resolve(),
    commonjs(),
    typescript({
      typescript: require("typescript"),
    }),
  ],
};
