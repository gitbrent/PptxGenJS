---
id: compatibility
title: Universal Compatibility
---

PptxGenJS works seamlessly in **modern web and Node environments**, thanks to dual ESM and CJS builds and zero runtime dependencies. Whether you're building a CLI tool, an Electron app, or a web-based presentation builder, the library adapts automatically to your stack.

### Supported Platforms

- **Node.js** – generate presentations in backend scripts, APIs, or CLI tools
- **React / Angular / Vite / Webpack** – just import and go, no config required
- **Electron** – build native apps with full filesystem access and PowerPoint output
- **Browser (Vanilla JS)** – embed in web apps with direct download support
- **Serverless / Edge Functions** – use in AWS Lambda, Vercel, Cloudflare Workers, etc.

> _Vite, Webpack, and modern bundlers automatically select the right build via the `exports` field in `package.json`._
> **Tip:** if you’re unsure, start with the **ES module build** (`pptxgen.es.js`).
> All modern bundlers and runtimes understand it, and it tree-shakes out the Node-only code paths automatically.

### Builds Provided

- **CommonJS**: [`dist/pptxgen.cjs.js`](https://github.com/gitbrent/PptxGenJS/dist/pptxgen.cjs.js)
- **ES Module**: [`dist/pptxgen.es.js`](https://github.com/gitbrent/PptxGenJS/dist/pptxgen.es.js)
