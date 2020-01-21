<h1 align="center">PptxGenJS</h1>
<h5 align="center">
  Create JavaScript PowerPoint Presentations
</h5>
<p align="center">
  <a href="https://github.com/gitbrent/PptxGenJS/">
    <img alt="PptxGenJS Sample Slides" title="PptxGenJS Sample Slides" src="https://raw.githubusercontent.com/gitbrent/PptxGenJS/gh-pages/img/readme_banner.png"/>
  </a>
</p>
<br/>

[![Known Vulnerabilities](https://snyk.io/test/npm/pptxgenjs/badge.svg)](https://snyk.io/test/npm/pptxgenjs) [![npm downloads](https://img.shields.io/npm/dm/pptxgenjs.svg)](https://www.npmjs.com/package/pptxgenjs) [![jsdelivr downloads](https://data.jsdelivr.com/v1/package/gh/gitbrent/pptxgenjs/badge)](https://www.jsdelivr.com/package/gh/gitbrent/pptxgenjs) [![typescripts definitions](https://img.shields.io/npm/types/pptxgenjs)](https://img.shields.io/npm/types/pptxgenjs)

# Table of Contents

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->


- [Introduction](#introduction)
- [Features](#features)
- [Live Demo](#live-demo)
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
- [Contributors](#contributors)
- [License](#license)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

# Introduction

This library creates Open Office XML (OOXML) Presentations which are compatible with Microsoft PowerPoint, Apple Keynote, and other applications.

# Features

**Extensive browser support**

- Create/download presentations on all current desktop & mobile web browsers
- IE11 is supported via bundle polyfill

**Major PowerPoint object types**

- Slides can include Charts, Images, Media, Shapes, Tables, Text and more.
- SVG images and YouTube videos are supported when viewed in PowerPoint online/2019+ desktop application

**Modern architecture**

- Supports client web browsers, NodeJS, and React/Angular/Electron
- Export methods return promises
- Client browsers have only a single dependency: JSZip
- Easy Angular/React integration (available via npm, cjs or es files)
- Typescript definitions included

**HTML to PowerPoint**

- Includes powerful [HTML-to-PowerPoint](#html-to-powerpoint-feature) feature to transform HTML tables into presentations with a single line of code

# Live Demo

Use the online demo to create a simple presentation to see how easy it is to use pptxgenjs, or check out the complete demo which showcases every available feature.

- [Simple Demo](https://gitbrent.github.io/PptxGenJS)
- [Complete Feature Demo](https://gitbrent.github.io/PptxGenJS/demo/)
- [PptxGenJS jsFiddle](https://jsfiddle.net/gitbrent/L1uctxm0/)

# Installation

## CDN

[jsDelivr Home](https://www.jsdelivr.com/package/gh/gitbrent/pptxgenjs)

Bundle: Modern Browsers and IE11
```html
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.0.0/dist/pptxgen.bundle.js"></script>
```

Min files: Modern Browsers
```html
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.0.0/libs/jszip.min.js"></script>
<script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.0.0/dist/pptxgen.min.js"></script>
```

## Download

[GitHub Latest Release](https://github.com/gitbrent/PptxGenJS/releases/latest)

Bundle: Modern Browsers and IE11
```html
<script src="PptxGenJS/dist/pptxgen.bundle.js"></script>
```

Min files: Modern Browsers
```html
<script src="PptxGenJS/libs/jszip.min.js"></script>
<script src="PptxGenJS/dist/pptxgen.min.js"></script>
```

## Npm

[PptxGenJS NPM Home](https://www.npmjs.com/package/pptxgenjs)

```bash
npm install pptxgenjs --save
```

## Yarn

```bash
yarn add pptxgenjs
```

## Additional Builds

- CommonJS: `dist/pptxgen.cjs.js`
- ES Module: `dist/pptxgen.es.js`

---

# Documentation

## Quick Start Guide

PptxGenJS PowerPoint presentations are created via JavaScript by following 4 basic steps:

1. Create a new Presentation
2. Add a Slide
3. Add one or more objects (Tables, Shapes, Images, Text and Media) to the Slide
4. Save the Presentation

```javascript
let pptx = new PptxGenJS();

let slide = pptx.addSlide();

let textboxText = "Hello World from PptxGenJS!";
let textboxOpts = { x: 1, y: 1, align: "center", color: "363636", fill: "f1f1f1" };
slide.addText(textboxText, textboxOpts);

pptx.writeFile("Sample Presentation");
```

That's really all there is to it!

---

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
- [Integration with Other Libraries](https://gitbrent.github.io/PptxGenJS/docs/integration.html)

---

## HTML-to-PowerPoint Feature

Easily convert HTML tables to PowerPoint presentations in a single call.

```javascript
let pptx = new PptxGenJS();
pptx.tableToSlides("tableElementId");
pptx.writeFile("HTML2PPT.pptx");
```

Learn more:

- [HTML-to-PowerPoint Documentation](https://gitbrent.github.io/PptxGenJS/docs/html-to-powerpoint.html)
- [Online HTML-to-PowerPoint Demo](https://gitbrent.github.io/PptxGenJS/demo/#html2pptx)

---

# Issues / Suggestions

Please file issues or suggestions on the [issues page on github](https://github.com/gitbrent/PptxGenJS/issues/new), or even better, [submit a pull request](https://github.com/gitbrent/PptxGenJS/pulls). Feedback is always welcome!

When reporting issues, please include a code snippet or a link demonstrating the problem.
Here is a small [jsFiddle](https://jsfiddle.net/gitbrent/L1uctxm0/) that is already configured and uses the latest PptxGenJS code.

---

# Need Help?

Sometimes implementing a new library can be a difficult task and the slightest mistake will keep something from working. We've all been there!

If you are having issues getting a presentation to generate, check out the code in the `demos` directory. There
are demos for both client browsers, node and react that contain working examples of every available library feature.

- Use a pre-configured jsFiddle to test with: [PptxGenJS Fiddle](https://jsfiddle.net/gitbrent/L1uctxm0/)
- [View questions tagged `PptxGenJS` on StackOverflow](https://stackoverflow.com/questions/tagged/pptxgenjs?sort=votes&pageSize=50). If you can't find your question, [ask it yourself](https://stackoverflow.com/questions/ask?tags=PptxGenJS) - be sure to tag it `PptxGenJS`.

---

# Contributors

Thank you to everyone for the issues, contributions and suggestions! ❤️

Special Thanks:

- [Dzmitry Dulko](https://github.com/DzmitryDulko) - Getting the project published on NPM
- [Michal Kacerovský](https://github.com/kajda90) - New Master Slide Layouts and Chart expertise
- [Connor Bowman](https://github.com/conbow) - Adding Placeholders
- [Reima Frgos](https://github.com/ReimaFrgos) - Multiple chart and general functionality patches
- [Matt King](https://github.com/kyrrigle) - Chart expertise
- [Mike Wilcox](https://github.com/clubajax) - Chart expertise

PowerPoint shape definitions and some XML code via [Officegen Project](https://github.com/Ziv-Barber/officegen)

---

# License

Copyright &copy; 2015-2020 [Brent Ely](https://github.com/gitbrent/PptxGenJS)

[MIT](https://github.com/gitbrent/PptxGenJS/blob/master/LICENSE)
