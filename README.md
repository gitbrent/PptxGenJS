[![MIT Licence](https://img.shields.io/github/license/gitbrent/pptxgenjs.svg)](https://opensource.org/licenses/mit-license.php)  [![Dependency Status](https://david-dm.org/gitbrent/pptxgenjs/status.svg)](https://david-dm.org/gitbrent/pptxgenjs)  [![Known Vulnerabilities](https://snyk.io/test/npm/pptxgenjs/badge.svg)](https://snyk.io/test/npm/pptxgenjs)  [![Package Quality](http://npm.packagequality.com/shield/pptxgenjs.png?style=flat-square)](https://github.com/gitbrent/pptxgenjs)  [![npm downloads](https://img.shields.io/npm/dm/pptxgenjs.svg)](https://www.npmjs.com/package/pptxgenjs)  [![jsdelivr downloads](https://data.jsdelivr.com/v1/package/gh/gitbrent/pptxgenjs/badge)](https://www.jsdelivr.com/package/gh/gitbrent/pptxgenjs)

# PptxGenJS

## JavaScript library that creates PowerPoint presentations

Quickly and easily create PowerPoint presentations with a few simple JavaScript commands in client web browsers or Node desktop apps.

### Main Features
* Widely Supported: Creates and downloads presentations on all current web browsers (Chrome, Edge, Firefox, etc.) and IE11
* Full Featured: Slides can include Charts, Images, Media, Shapes, Tables and Text (plus Master Slides/Templates)
* Easy To Use: Entire PowerPoint presentations can be created in a few lines of code
* Modern: Pure JavaScript solution - everything necessary to create PowerPoint PPT exports is included

### Additional Features
* Use the unique [HTML-to-PowerPoint](#html-to-powerpoint-feature) feature to copy an HTML table into 1 or more Slides with a single command

**************************************************************************************************

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->
**Table of Contents**  (*generated with [DocToc](https://github.com/thlorenz/doctoc)*)

- [Live Demo](#live-demo)
  - [Installation](#installation)
    - [CDN](#cdn)
    - [Download](#download)
    - [Npm](#npm)
    - [Yarn](#yarn)
- [Quick Start Guide](#quick-start-guide)
- [Library API](#library-api)
  - [Presentation Creation/Options](#presentation-creationoptions)
  - [Slide Creation/Options](#slide-creationoptions)
  - [Saving a Presentation](#saving-a-presentation)
  - [Master Slides and Corporate Branding](#master-slides-and-corporate-branding)
  - [Adding Charts](#adding-charts)
  - [Adding Images](#adding-images)
  - [Adding Media (Audio/Video/YouTube)](#adding-media-audiovideoyoutube)
  - [Adding Shapes](#adding-shapes)
  - [Adding Tables](#adding-tables)
  - [Adding Text](#adding-text)
  - [Including Speaker Notes](#including-speaker-notes)
  - [Using Scheme Colors](#using-scheme-colors)
- [HTML-to-PowerPoint Feature](#html-to-powerpoint-feature)
- [Integration with Other Libraries](#integration-with-other-libraries)
- [Full PowerPoint Shape Library](#full-powerpoint-shape-library)
- [Typescript Definitions](#typescript-definitions)
- [Issues / Suggestions](#issues--suggestions)
- [Need Help?](#need-help)
- [Unimplemented Features](#unimplemented-features)
- [Coming Soon ⏰](#coming-soon-)
- [Contributors ✨](#contributors-)
- [License](#license)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

**************************************************************************************************
# Live Demo
Use JavaScript to create a PowerPoint presentation with your web browser right now:  
[https://gitbrent.github.io/PptxGenJS](https://gitbrent.github.io/PptxGenJS)

## Installation

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

**************************************************************************************************
# Quick Start Guide
PptxGenJS PowerPoint presentations are created via JavaScript by following 4 basic steps:

1. Create a new Presentation
2. Add a Slide
3. Add one or more objects (Tables, Shapes, Images, Text and Media) to the Slide
4. Save the Presentation

```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();
slide.addText('Hello World!', { x:1.5, y:1.5, fontSize:18, color:'363636' });
pptx.save('Sample Presentation');
```
That's really all there is to it!


**************************************************************************************************
# Library API

## Presentation Creation/Options
[Creating a Presentation](https://gitbrent.github.io/PptxGenJS/docs/usage-pres-create.html)  
[Presentation Options](https://gitbrent.github.io/PptxGenJS/docs/usage-pres-options.html)  

## Slide Creation/Options
[Adding a Slide](https://gitbrent.github.io/PptxGenJS/docs/usage-add-slide.html)  
[Slide Options](https://gitbrent.github.io/PptxGenJS/docs/usage-slide-options.html)

## Saving a Presentation
[Saving a Presentation](https://gitbrent.github.io/PptxGenJS/docs/usage-saving.html)

## Master Slides and Corporate Branding
[Master Slides](https://gitbrent.github.io/PptxGenJS/docs/masters.html)

## Adding Charts
[Adding Charts](https://gitbrent.github.io/PptxGenJS/docs/api-charts.html)

## Adding Images
[Adding Images](https://gitbrent.github.io/PptxGenJS/docs/api-images.html)

## Adding Media (Audio/Video/YouTube)
[Adding Media](https://gitbrent.github.io/PptxGenJS/docs/api-media.html)

## Adding Shapes
[Adding Shapes](https://gitbrent.github.io/PptxGenJS/docs/api-shapes.html)

## Adding Tables
[Adding Tables](https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)

## Adding Text
[Adding Text](https://gitbrent.github.io/PptxGenJS/docs/api-text.html)

## Including Speaker Notes
[Speaker Notes](https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html)

## Using Scheme Colors
[Using Scheme Colors](https://gitbrent.github.io/PptxGenJS/docs/shapes-and-schemes.html)


**************************************************************************************************
# HTML-to-PowerPoint Feature

[HTML-to-PowerPoint](https://gitbrent.github.io/PptxGenJS/docs/html-to-powerpoint.html)


**************************************************************************************************
# Integration with Other Libraries

[Integration with Other Libraries](https://gitbrent.github.io/PptxGenJS/docs/integration.html)


**************************************************************************************************
# Full PowerPoint Shape Library
If you are planning on creating Shapes (basically anything other than Text, Tables or Rectangles), then you'll want to
include the `pptxgen.shapes.js` library.

The shapes file contains a complete PowerPoint Shape object array thanks to the [officegen project](https://github.com/Ziv-Barber/officegen).

```javascript
<script src="PptxGenJS/dist/pptxgen.shapes.js"></script>
```


**************************************************************************************************
# Typescript Definitions

As of version 2.3.0, typescript definitions are available (`pptxgen.d.ts`).


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
# Coming Soon ⏰

The library is currently being rewritten in TypeScript for version 3.0 which will be completed by the end of 2019.  

Visit the "version-3.0" branch to try it out
* [PptxGenJS 3.0 Preview](https://github.com/gitbrent/PptxGenJS/tree/version-3.0)

New Features
* Brand-new TypeScript/ES6 Class codebase eliminated dozens of bugs and greatly increased stability
* Code is logically separated into 10+ files, making pull requests and maintenance easier
* Completely rewritten Table AutoPaging and HTML-to-PowerPoint methods - faster and much more accurate
* Save/Export:
 * Promise-based export methods - no more callbacks
 * Two new methods (Write and WriteFile) will replace `save()`
 * Supports all types of output methods: ArrayBuffer, Blob, etc.

Outstanding Dev Items
* Angular/React integration has not been completed as of yet (but it will be MUCH EASIER once it is finalized)
* TypeScript definitions are not up-to-date
* SlideNumbers do not work
* `save()` is still the only export method: `write()` & `writeFile()` are coming in September
* Correct MIME type for zip exports
* Other small items

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
