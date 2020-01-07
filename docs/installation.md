---
id: installation
title: Installation
---

## Installation Notes
* Dependencies: `JSZip`
* Bundle Includes: `jszip.min.js`, and `promise.min.js`
* IE11 Support: The `promise.min.js` promise polyfill is **required**

## Installation Methods

### CDN
```html
<!-- Bundle: Easiest to use, supports all browsers -->
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.0.0/dist/pptxgen.bundle.js"></script>

<!-- Individual files: Add only what's needed to avoid clobbering loaded libraries -->
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.0.0/libs/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.0.0/dist/pptxgen.min.js"></script>
```

### Download
[GitHub Latest Release](https://github.com/gitbrent/PptxGenJS/releases/latest)
```html
<!-- Bundle: Easiest to use, supports all browsers -->
<script src="PptxGenJS/dist/pptxgen.bundle.js"></script>

<!-- Individual files: Add only what's needed to avoid clobbering loaded libraries -->
<script src="PptxGenJS/libs/jszip.min.js"></script>
<script src="PptxGenJS/dist/pptxgen.min.js"></script>
<!-- IE11 requires Promises polyfill -->
<!-- <script src="PptxGenJS/libs/promise.min.js"></script> -->
```

## Npm
[PptxGenJS NPM Home](https://www.npmjs.com/package/pptxgenjs)
```bash
npm install pptxgenjs --save
```
```javascript
let PptxGenJS = require("pptxgenjs");
let pptx = new PptxGenJS();
```

## Yarn
```bash
yarn add pptxgenjs
```

## Additional Builds
* CommonJS: `dist/pptxgen.cjs.js`
* ES Module: `dist/pptxgen.es.js`
