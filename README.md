
[![Package Quality](http://npm.packagequality.com/shield/pptxgenjs.png?style=flat-square)](https://github.com/gitbrent/pptxgenjs)  [![Dependency Status](https://david-dm.org/gitbrent/pptxgenjs/status.svg)](https://david-dm.org/gitbrent/pptxgenjs)  [![Known Vulnerabilities](https://snyk.io/test/npm/pptxgenjs/badge.svg)](https://snyk.io/test/npm/pptxgenjs)  [![npm downloads](https://img.shields.io/npm/dm/pptxgenjs.svg)](https://www.npmjs.com/package/pptxgenjs)  [![jsdelivr downloads](https://data.jsdelivr.com/v1/package/gh/gitbrent/pptxgenjs/badge)](https://www.jsdelivr.com/package/gh/gitbrent/pptxgenjs)  [![typescripts definitions](https://img.shields.io/npm/types/pptxgenjs)](https://img.shields.io/npm/types/pptxgenjs)

# PptxGenJS

JavaScript library that creates PowerPoint presentations
* Creates presentations on all current web browsers and IE11
* Slides can include Charts, Images, Media, Shapes, Tables and Text, etc.
* Powerful [HTML-to-PowerPoint](#html-to-powerpoint-feature) feature to transform any HTML table into a presentation
* Modern, pure JavaScript, promise-based library
* Only a single dependency (JSZip)
* Easy Angular/React integration (available via npm, cjs or es files)

**************************************************************************************************

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->
**Table of Contents**

- [Demo](#demo)
- [Installation](#installation)
  - [CDN](#cdn)
  - [Download](#download)
  - [Npm](#npm)
  - [Yarn](#yarn)
  - [Additional Builds](#additional-builds)
- [Documentation](#documentation)
  - [Quick Start Guide](#quick-start-guide)
  - [Library API](#library-api)
  - [HTML-to-PowerPoint Feature](#html-to-powerpoint-feature)
- [Issues / Suggestions](#issues--suggestions)
- [Need Help?](#need-help)
- [Unimplemented Features](#unimplemented-features)
- [Contributors ✨](#contributors-)
- [License](#license)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

**************************************************************************************************
# Demo
Use JavaScript to create a PowerPoint presentation with your web browser right now!  
* [https://gitbrent.github.io/PptxGenJS](https://gitbrent.github.io/PptxGenJS)

The complete library demo is also online.
* [https://gitbrent.github.io/PptxGenJS/demo/](https://gitbrent.github.io/PptxGenJS/demo/)


# Installation

## CDN
```html
<!-- Bundle: Easiest to use, supports all browsers -->
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.0.0/dist/pptxgen.bundle.js"></script>

<!-- Individual files: Add only what's needed to avoid clobbering loaded libraries -->
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.0.0/libs/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.0.0/dist/pptxgen.min.js"></script>
```

## Download
[GitHub Latest Release](https://github.com/gitbrent/PptxGenJS/releases/latest)
```html
<!-- Bundle: Easiest to use, supports all browsers -->
<script src="PptxGenJS/libs/pptxgen.bundle.js"></script>

<!-- Individual files: Add only what's needed to avoid clobbering loaded libraries -->
<script src="PptxGenJS/libs/jszip.min.js"></script>
<script src="PptxGenJS/dist/pptxgen.min.js"></script>
<!-- <script src="PptxGenJS/libs/promise.min.js"></script> IE11 requires Promises polyfill -->
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
* CommonJS: `dist/pptxgenjs.cjs.js`
* ES Module: `dist/pptxgenjs.es.js`

**************************************************************************************************

# Documentation

## Quick Start Guide
PptxGenJS PowerPoint presentations are created via JavaScript by following 4 basic steps:

1. Create a new Presentation
2. Add a Slide
3. Add one or more objects (Tables, Shapes, Images, Text and Media) to the Slide
4. Save the Presentation

```javascript
var pptx = new PptxGenJS();
var slide = pptx.addSlide();
slide.addText(
  'Hello World from PptxGenJS!',
  { x:1, y:1, w:'80%', h:3, color:'363636', align:'center', fill:'f1f1f1' }
);
pptx.writeFile('Sample Presentation');
```
That's really all there is to it!


**************************************************************************************************
## Library API
Full documentation and code examples are available
- [Creating a Presentation](https://gitbrent.github.io/PptxGenJS/docs/usage-pres-create.html)  
- [Presentation Options](https://gitbrent.github.io/PptxGenJS/docs/usage-pres-options.html)  
- [Adding a Slide](https://gitbrent.github.io/PptxGenJS/docs/usage-add-slide.html)  
- [Slide Options](https://gitbrent.github.io/PptxGenJS/docs/usage-slide-options.html)
- [Saving a Presentation](https://gitbrent.github.io/PptxGenJS/docs/usage-saving.html)
- [Master Slides](https://gitbrent.github.io/PptxGenJS/docs/masters.html)
- [Adding Charts](https://gitbrent.github.io/PptxGenJS/docs/api-charts.html)
- [Adding Images](https://gitbrent.github.io/PptxGenJS/docs/api-images.html)
- [Adding Media](https://gitbrent.github.io/PptxGenJS/docs/api-media.html)
- [Adding Shapes](https://gitbrent.github.io/PptxGenJS/docs/api-shapes.html)
- [Adding Tables](https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)
- [Adding Text](https://gitbrent.github.io/PptxGenJS/docs/api-text.html)
- [Speaker Notes](https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html)
- [Using Scheme Colors](https://gitbrent.github.io/PptxGenJS/docs/shapes-and-schemes.html)
- [Creating a Presentation](https://gitbrent.github.io/PptxGenJS/docs/installation.html)  
- [Integration with Other Libraries](https://gitbrent.github.io/PptxGenJS/docs/integration.html)

Note: Typescript Definitions are included


**************************************************************************************************
## HTML-to-PowerPoint Feature
Easily convert HTML tables to PowerPoint presentations in a single call.

```javascript
var pptx = new PptxGenJS();
pptx.tableToSlides('tableId');
pptx.writeFile('HTML-table.pptx');
```

Learn more:
- [HTML-to-PowerPoint Documentation](https://gitbrent.github.io/PptxGenJS/docs/html-to-powerpoint.html)
- [Online HTML-to-PowerPoint Demo](https://gitbrent.github.io/PptxGenJS/demo/#tab2)


**************************************************************************************************
# Issues / Suggestions

Please file issues or suggestions on the [issues page on github](https://github.com/gitbrent/PptxGenJS/issues/new), or even better, [submit a pull request](https://github.com/gitbrent/PptxGenJS/pulls). Feedback is always welcome!

When reporting issues, please include a code snippet or a link demonstrating the problem.
Here is a small [jsFiddle](https://jsfiddle.net/gitbrent/gx34jy59/5/) that is already configured and uses the latest PptxGenJS code.


**************************************************************************************************
# Need Help?

Sometimes implementing a new library can be a difficult task and the slightest mistake will keep something from working. We've all been there!

If you are having issues getting a presentation to generate, check out the demos in the `examples` directory. There
are demos for both Nodejs and client-browsers that contain working examples of every available library feature.

* Use a pre-configured jsFiddle to test with: [PptxGenJS Fiddle](https://jsfiddle.net/gitbrent/gx34jy59/)
* [View questions tagged `PptxGenJS` on StackOverflow](https://stackoverflow.com/questions/tagged/pptxgenjs?sort=votes&pageSize=50).  If you can't find your question, [ask it yourself](https://stackoverflow.com/questions/ask?tags=PptxGenJS) - be sure to tag it `PptxGenJS`.


**************************************************************************************************
# Unimplemented Features

The PptxGenJS library is not designed to replicate all the functionality of PowerPoint, meaning several features
are not on the development roadmap.

These include:
* Animations
* Importing Existing Presentations and/or Templates
* Outlines
* SmartArt


**************************************************************************************************
# Contributors ✨

Thank you to everyone for the issues, contributions and suggestions! ❤️

Special Thanks:
* [Dzmitry Dulko](https://github.com/DzmitryDulko) - Getting the project published on NPM
* [Michal Kacerovský](https://github.com/kajda90) - New Master Slide Layouts and Chart expertise
* [Connor Bowman](https://github.com/conbow) - Adding Placeholders
* [Reima Frgos](https://github.com/ReimaFrgos) - Multiple chart and general functionality patches
* [Matt King](https://github.com/kyrrigle) - Chart expertise
* [Mike Wilcox](https://github.com/clubajax) - Chart expertise

PowerPoint shape definitions and some XML code via [Officegen Project](https://github.com/Ziv-Barber/officegen)

**************************************************************************************************
# License

Copyright &copy; 2015-2019 [Brent Ely](https://github.com/gitbrent/PptxGenJS)

[MIT](https://github.com/gitbrent/PptxGenJS/blob/master/LICENSE)
