[![npm version](https://img.shields.io/npm/v/pptxgenjs.svg)](https://www.npmjs.com/package/pptxgenjs)  [![MIT Licence](https://img.shields.io/github/license/gitbrent/pptxgenjs.svg)](https://opensource.org/licenses/mit-license.php)  [![npm downloads](https://img.shields.io/npm/dm/pptxgenjs.svg)](https://www.npmjs.com/package/pptxgenjs)  [![Package Quality](http://npm.packagequality.com/shield/pptxgenjs.png?style=flat-square)](https://github.com/gitbrent/pptxgenjs)

# PptxGenJS

### JavaScript library that produces PowerPoint (pptx) presentations

Quickly and easily create PowerPoint presentations with a few simple JavaScript commands in client web browsers or Node desktop apps.

## Main Features
* Widely Supported: Creates and downloads presentations on all current web browsers (Chrome, Edge, Firefox, etc.) and IE11
* Full Featured: Slides can include Charts, Images, Media, Shapes, Tables and Text (plus Master Slides/Templates)
* Easy To Use: Entire PowerPoint presentations can be created in a few lines of code
* Modern: Pure JavaScript solution - everything necessary to create PowerPoint PPT exports is included

## Additional Features
* Use the unique [HTML-to-PowerPoint](#html-to-powerpoint-feature) feature to copy an HTML table into 1 or more Slides with a single command

**************************************************************************************************

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->
**Table of Contents**  (*generated with [DocToc](https://github.com/thlorenz/doctoc)*)

- [Live Demo](#live-demo)
- [Installation](#installation)
  - [Client-Side](#client-side)
    - [Include Local Scripts](#include-local-scripts)
    - [Include Bundled Script](#include-bundled-script)
    - [Install With Bower](#install-with-bower)
  - [Node.js](#nodejs)
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
- [Special Thanks](#special-thanks)
- [Support Us](#support-us)
- [License](#license)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

**************************************************************************************************
# Live Demo
Use JavaScript to create a PowerPoint presentation with your web browser right now:
[https://gitbrent.github.io/PptxGenJS](https://gitbrent.github.io/PptxGenJS)

# Installation
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
* Desktop: Compatible with Electron applications

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
* Use Ask Question on [StackOverflow](http://stackoverflow.com/) - be sure to tag it with "PptxGenJS"


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
# Special Thanks

* [Officegen Project](https://github.com/Ziv-Barber/officegen) - Shape definitions and XML code
* [Dzmitry Dulko](https://github.com/DzmitryDulko) - Getting the project published on NPM
* [kajda90](https://github.com/kajda90) - New Master Slide Layouts
* [Connor Bowman](https://github.com/conbow) - Adding Placeholders
* PPTX Chart Experts: [kajda90](https://github.com/kajda90), [Matt King](https://github.com/kyrrigle), [Mike Wilcox](https://github.com/clubajax)
* Everyone who has [contributed](https://github.com/gitbrent/PptxGenJS/graphs/contributors), submitted an Issue, or created Pull Request.


**************************************************************************************************
# Support Us

Do you like this library and find it useful?  Tell the world about us! [PptxGenJS project](https://github.com/gitbrent/PptxGenJS)

Thanks to everyone who supports this project! &#10084;


**************************************************************************************************
# License

Copyright &copy; 2015-2018 [Brent Ely](https://github.com/gitbrent/PptxGenJS)

[MIT](https://github.com/gitbrent/PptxGenJS/blob/master/LICENSE)
