---
id: installation
title: Installation
---

## Installation Notes
* Dependencies: `jQuery` and `jsZip`
* Bundle Includes: `jquery.min.js`, `jszip.min.js`, and `promise.min.js`
* IE11 Support: The `promise.min.js` promise polyfill is **required**

## Installation Methods

### CDN
```html
<!-- Bundle: Easiest to use, supports all browsers -->
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@2.5.0/dist/pptxgen.bundle.js"></script>

<!-- Individual files: Add only what's needed to avoid clobbering loaded libraries -->
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@2.5.0/libs/jquery.min.js"></script>
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@2.5.0/libs/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@2.5.0/dist/pptxgen.min.js"></script>
```

### Download
[GitHub Latest Release](https://github.com/gitbrent/PptxGenJS/releases/latest)
```html
<!-- Bundle: Easiest to use, supports all browsers -->
<script src="PptxGenJS/libs/pptxgen.bundle.js"></script>

<!-- Individual files: Add only what's needed to avoid clobbering loaded libraries -->
<script src="PptxGenJS/libs/jquery.min.js"></script>
<script src="PptxGenJS/libs/jszip.min.js"></script>
<script src="PptxGenJS/dist/pptxgen.min.js"></script>
<!-- IE11 requires Promises polyfill -->
<!-- <script src="PptxGenJS/libs/promise.min.js"></script> -->
```

### Npm
[PptxGenJS NPM Home](https://www.npmjs.com/package/pptxgenjs)
```javascript
npm install pptxgenjs

var pptx = require("pptxgenjs");
```

### Yarn
```ksh
yarn install pptxgenjs
```
