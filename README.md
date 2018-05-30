[![Open Source Love](https://badges.frapsoft.com/os/v1/open-source.svg?v=103)](https://github.com/ellerbrock/open-source-badge/)  [![MIT Licence](https://img.shields.io/github/license/gitbrent/pptxgenjs.svg)](https://opensource.org/licenses/mit-license.php)  [![npm version](https://img.shields.io/npm/v/pptxgenjs.svg)](https://www.npmjs.com/package/pptxgenjs)  [![npm downloads](https://img.shields.io/npm/dm/pptxgenjs.svg)](https://www.npmjs.com/package/pptxgenjs)

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

- [Documentation](#documentation)
- [Live Demo](#live-demo)
- [Installation](#installation)
  - [Client-Side](#client-side)
    - [Include Local Scripts](#include-local-scripts)
    - [Include Bundled Script](#include-bundled-script)
    - [Install With Bower](#install-with-bower)
  - [Node.js](#nodejs)
- [Presentations: Usage and Options](#presentations-usage-and-options)
  - [Presentation Creation/Options](#presentation-creationoptions)
  - [Slide Creation/Options](#slide-creationoptions)
  - [Saving a Presentation](#saving-a-presentation)
  - [Adding Charts](#adding-charts)
  - [Adding Text](#adding-text)
  - [Adding Tables](#adding-tables)
  - [Adding Shapes](#adding-shapes)
  - [Adding Images](#adding-images)
  - [Adding Media (Audio/Video/YouTube)](#adding-media-audiovideoyoutube)
- [Master Slides and Corporate Branding](#master-slides-and-corporate-branding)
- [HTML-to-PowerPoint Feature](#html-to-powerpoint-feature)
- [Scheme Colors](#scheme-colors)
- [Integration with Other Libraries](#integration-with-other-libraries)
- [Full PowerPoint Shape Library](#full-powerpoint-shape-library)
- [Issues / Suggestions](#issues--suggestions)
- [Need Help?](#need-help)
- [Version 2.0 Breaking Changes](#version-20-breaking-changes)
  - [All Users](#all-users)
  - [Node Users](#node-users)
- [Unimplemented Features](#unimplemented-features)
- [Special Thanks](#special-thanks)
- [Support Us](#support-us)
- [License](#license)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

**************************************************************************************************
# Documentation

There's more than just the README!  
* View the online [API Reference](https://gitbrent.github.io/PptxGenJS/docs/installation.html)


**************************************************************************************************
# Live Demo
Use JavaScript to Create a PowerPoint presentation with your web browser right now:
[https://gitbrent.github.io/PptxGenJS](https://gitbrent.github.io/PptxGenJS)

# Installation
## Client-Side
### Include Local Scripts
```javascript
<script lang="javascript" src="PptxGenJS/libs/jquery.min.js"></script>
<script lang="javascript" src="PptxGenJS/libs/jszip.min.js"></script>
<script lang="javascript" src="PptxGenJS/dist/pptxgen.js"></script>
```
* IE11 support requires a Promises polyfill as well (included in the libs folder)

### Include Bundled Script
```javascript
<script lang="javascript" src="PptxGenJS/dist/pptxgen.bundle.js"></script>
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
# Presentations: Usage and Options
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
## Presentation Creation/Options

[Creating a Presentation](https://gitbrent.github.io/PptxGenJS/docs/usage-pres-create.html)

[Presentation Options](https://gitbrent.github.io/PptxGenJS/docs/usage-pres-options.html)


**************************************************************************************************
## Slide Creation/Options

[Adding a Slide](https://gitbrent.github.io/PptxGenJS/docs/usage-add-slide.html)

[Slide Options](https://gitbrent.github.io/PptxGenJS/docs/usage-slide-options.html)


**************************************************************************************************
## Saving a Presentation

[Saving a Presentation](https://gitbrent.github.io/PptxGenJS/docs/usage-saving.html)


**************************************************************************************************
## Adding Charts

[Adding Charts](https://gitbrent.github.io/PptxGenJS/docs/api-charts.html)


**************************************************************************************************
## Adding Text

[Adding Text](https://gitbrent.github.io/PptxGenJS/docs/api-text.html)


**************************************************************************************************
## Adding Tables

[Adding Tables](https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)


**************************************************************************************************
## Adding Shapes

[Adding Shapes](https://gitbrent.github.io/PptxGenJS/docs/api-shapes.html)


**************************************************************************************************
## Adding Images

[Adding Images](https://gitbrent.github.io/PptxGenJS/docs/api-images.html)


**************************************************************************************************
## Adding Media (Audio/Video/YouTube)

[Adding Media](https://gitbrent.github.io/PptxGenJS/docs/api-media.html)


**************************************************************************************************
# Master Slides and Corporate Branding

[Master Slides](https://gitbrent.github.io/PptxGenJS/docs/masters.html)


**************************************************************************************************
# HTML-to-PowerPoint Feature

[HTML-to-PowerPoint](https://gitbrent.github.io/PptxGenJS/docs/html-to-powerpoint.html)


**************************************************************************************************
# Scheme Colors

[Scheme Colors](https://gitbrent.github.io/PptxGenJS/docs/shapes-and-schemes.html)


**************************************************************************************************
# Integration with Other Libraries

[Integration with Other Libraries](https://gitbrent.github.io/PptxGenJS/docs/integration.html)


**************************************************************************************************
# Full PowerPoint Shape Library
If you are planning on creating Shapes (basically anything other than Text, Tables or Rectangles), then you'll want to
include the `pptxgen.shapes.js` library.

The shapes file contains a complete PowerPoint Shape object array thanks to the [officegen project](https://github.com/Ziv-Barber/officegen).

```javascript
<script lang="javascript" src="PptxGenJS/dist/pptxgen.shapes.js"></script>
```


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

* Use a pre-configured jsFiddle to test with: [PptxGenJS Fiddle](https://jsfiddle.net/gitbrent/gx34jy59/5/)
* Use Ask Question on [StackOverflow](http://stackoverflow.com/) - be sure to tag it with "PptxGenJS"

**************************************************************************************************
# Version 2.0 Breaking Changes

Please note that version 2.0.0 enabled some much needed cleanup, but may break your previous code...
(however, a quick search-and-replace will fix any issues).

While the changes may only impact cosmetic properties, it's recommended you test your solutions thoroughly before upgrading PptxGenJS to the 2.0 version.

## All Users
The library `getVersion()` method is now a property: `version`

Option names are now caseCase across all methods:
* `font_face` renamed to `fontFace`
* `font_size` renamed to `fontSize`
* `line_dash` renamed to `lineDash`
* `line_head` renamed to `lineHead`
* `line_size` renamed to `lineSize`
* `line_tail` renamed to `lineTail`

Options deprecated in early 1.0 versions (hopefully nobody still uses these):
* `marginPt` renamed to `margin`


## Node Users

**Major Change**
* `require('pptxgenjs')` no longer returns a singleton instance
* `pptx = new PptxGenJS()` will create a single, unique instance
* Advantage: Creating [multiple presentations](#saving-multiple-presentations) is much easier now - see [Issue #83](https://github.com/gitbrent/PptxGenJS/issues/83) for more).

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

* [Officegen Project](https://github.com/Ziv-Barber/officegen) - For the Shape definitions and XML code
* [Dzmitry Dulko](https://github.com/DzmitryDulko) - For getting the project published on NPM
* [kajda90](https://github.com/kajda90) - For the new Master Slide Layouts
* PPTX Chart Experts: [kajda90](https://github.com/kajda90), [Matt King](https://github.com/kyrrigle), [Mike Wilcox](https://github.com/clubajax)
* Everyone who has submitted an Issue or Pull Request. :-)

**************************************************************************************************
# Support Us

Do you like this library and find it useful?  Add a link to the [PptxGenJS project](https://github.com/gitbrent/PptxGenJS)
on your blog, website or social media.

Thanks to everyone who supports this project! <3

**************************************************************************************************
# License

Copyright &copy; 2015-2018 [Brent Ely](https://github.com/gitbrent/PptxGenJS)

[MIT](https://github.com/gitbrent/PptxGenJS/blob/master/LICENSE)
