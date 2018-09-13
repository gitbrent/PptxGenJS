---
id: installation
title: Installation
---

## Client-Side
### Include Local Scripts
```javascript
<script src="PptxGenJS/libs/jquery.min.js"></script>
<script src="PptxGenJS/libs/jszip.min.js"></script>
<script src="PptxGenJS/dist/pptxgen.js"></script>
```
* IE11 support requires a Promises polyfill as well (included in the libs folder)

### Include Bundled Script
```javascript
<script src="PptxGenJS/dist/pptxgen.bundle.js"></script>
```
* Bundle script includes all libraries: jQuery + JSzip + PptxGenJS + Promises

### Use CDN
```javascript
<script src="https://cdn.rawgit.com/gitbrent/PptxGenJS/8cb0150f/dist/pptxgen.bundle.js"></script>
```

### Install With Bower
```javascript
bower install pptxgen
```

## Node.js
[PptxGenJS NPM Homepage](https://www.npmjs.com/package/pptxgenjs)
```javascript
npm install pptxgenjs

var pptx = require("pptxgenjs");
```
* Desktop: Compatible with Electron applications!
