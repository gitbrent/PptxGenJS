---
id: integration
title: Integration by Environment
---

PptxGenJS can be used in various JavaScript environments. Choose the integration method below that best suits your project setup.

## Available Distributions

- ES6 Module `dist/pptxgen.es.js`
- CommonJS `dist/pptxgen.cjs.js`
- Browser `dist/pptxgen.min.js`

## Environment Guide

| Environment(s)                                                                                                | Import / Usage                                                                                                         | Notes / Details                                                                                                                                                                                                          |
| ------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Node.js (Version 18 and higher)**| `import pptxgen from "pptxgenjs"`| Automatically uses the appropriate Node.js build based on your project's module type (`package.json#type`). Both ESM and CommonJS formats are fully supported.|
| **Browser Bundlers** (Webpack, Vite, Rollup, Parcel, Browserify, Create React App, Next.js, Angular, Vue CLI) | `import pptxgen from 'pptxgenjs'`| Your bundler will automatically select the optimized ES Module build (`dist/pptxgen.es.js`). This enables effective tree-shaking to minimize your final bundle size. No extra bundler configuration is typically needed. |
| **Plain Browser (`<script>` tag, no bundler)**| Include the bundled script directly in your HTML: `<script src=".../pptxgen.bundle.js"></script>`| This provides a self-contained build (`dist/pptxgen.bundle.js`) that adds the `PptxGenJS` object to the global `window` scope. Useful for simple scripts or environments without a module bundler.|
| **Web Worker / Service Worker**| `import pptxgen from 'pptxgenjs'` (Requires a module worker (`type: "module"`) or the use of import maps)| Utilize the ES Module build (`dist/pptxgen.es.js`). Remember that data (like the final presentation `ArrayBuffer`) will need to be transferred back to the main thread using `postMessage`.|
| **Serverless Functions** (AWS Lambda, Cloudflare Workers, etc.)| `import pptxgen from 'pptxgenjs'` (for ESM runtimes) OR `const pptxgen = require('pptxgenjs')` (for CommonJS runtimes) | Bundle your function code using a tool like esbuild or Vite SSR; Be mindful of function size limits and potential cold start impacts from larger dependencies.|
| **Electron (Main Process)**| Same as **Node.js**| In the main Electron process, you have full access to Node.js APIs, including the filesystem, which is useful for directly saving presentation files using the `writeFile()` method.|
| **Electron (Renderer Process)**| Same as **Browser Bundlers**| The renderer process is similar to a browser environment. If `nodeIntegration` is enabled and securely configured, you may also be able to use Node.js filesystem access from the renderer.|

## Integration Demos

Many of the common integration methods have working demos and code available.

### React + Vite

- Online Demo: [Demo Page](https://gitbrent.github.io/PptxGenJS/demo/vite/index.html)
- Source Code: [GitHub Repo](https://github.com/gitbrent/PptxGenJS/tree/master/demos/vite-demo)

### Node.js

- Source Code: [GitHub Repo](https://github.com/gitbrent/PptxGenJS/tree/master/demos/node)

### Web Browser (script)

- Online Demo: [Demo Page](https://gitbrent.github.io/PptxGenJS/demo/browser/index.html)
- Source Code: [GitHub Repo](https://github.com/gitbrent/PptxGenJS/tree/master/demos/browser)

### Web Worker

- Online Demo: [Demo Page](https://gitbrent.github.io/PptxGenJS/demo/browser/worker_test.html)
- Source Code: [GitHub Repo](https://github.com/gitbrent/PptxGenJS/tree/master/demos/browser)

## Troubleshooting

### Webpack

The pptxgenjs library has been tested with several different framework and bundler combinations
including Vite and Webpack. While most projects can simply install pptxgenjs and go, there are times
when errors occur.

Here's an example from the latest version of Docusaurus (v3.7) where Webpack (v5) fails during the build process:

```text
[ERROR] Client bundle compiled with errors therefore further build is impossible.
Module build failed: UnhandledSchemeError: Reading from "node:fs" is not handled by plugins (Unhandled scheme).
Webpack supports "data:" and "file:" URIs by default.
You may need an additional plugin to handle "node:" URIs.
```

The error is being caused by the use of the "node:" prefix within "browser" field in pptxgenjs' `package.json` file.

```json
// pptxgenjs/package.json
{
  "name": "pptxgenjs",
  "browser": {
    "express": false,
    "fs": false,
    "https": false,
    "image-size": false,
    "node:fs": false,
    "node:https": false
  }
}
```

Starting in version 4.0.0, node modules are excluded using this format in the pptxgenjs source code.

### Excluding "node:" from Webpack Builds

The fix for docusaurus involved creating a custom plug-in to tell webpack to ignore any "node:" references.

1. Add the [plug-in code](https://github.com/gitbrent/PptxGenJS/tree/gh-pages/docusaurus-plugin)
2. Add this to your package.json: `"pptxgenjs-plugin": "file:docusaurus-plugin/src"`
3. Run `npm i`
4. Add this to "docusaurus.config.js": `plugins: ["pptxgenjs-plugin"]`
5. Webpack now ignores "node:" items from the "browser" field and builds successfully
