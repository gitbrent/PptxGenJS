---
id: api0
title: This is document number 3
---
# PptxGenJS

### JavaScript library that produces PowerPoint (pptx) presentations

Quickly and easily create PowerPoint presentations with a few simple JavaScript commands in client web browsers or Node desktop apps.

## Main Features
* Widely Supported: Creates and downloads presentations on all current web browsers (Chrome, Edge, Firefox, etc.) and IE11
* Full Featured: Slides can include Charts, Images, Media, Shapes, Tables and Text (plus Master Slides/Templates)
* Easy To Use: Entire PowerPoint presentations can be created in a few lines of code
* Modern: Pure JavaScript solution - everything necessary to create PowerPoint PPT exports is included

## Additional Features
* Use the unique [Table-to-Slides](#table-to-slides-feature) feature to copy an HTML table into 1 or more Slides with a single command

**************************************************************************************************

**Table of Contents**  (*generated with [DocToc](https://github.com/thlorenz/doctoc)*)

- [Live Demo](#live-demo)
- [Installation](#installation)
  - [Client-Side](#client-side)
    - [Include Local Scripts](#include-local-scripts)
    - [Include Bundled Script](#include-bundled-script)
    - [Install With Bower](#install-with-bower)
  - [Node.js](#nodejs)
- [Presentations: Usage and Options](#presentations-usage-and-options)
  - [Creating a Presentation](#creating-a-presentation)
    - [Presentation Properties](#presentation-properties)
    - [Presentation Layouts](#presentation-layouts)
    - [Presentation Layout Options](#presentation-layout-options)
    - [Presentation Text Direction](#presentation-text-direction)
  - [Adding a Slide](#adding-a-slide)
    - [Slide Formatting](#slide-formatting)
    - [Slide Formatting Options](#slide-formatting-options)
    - [Applying Master Slides / Branding](#applying-master-slides--branding)
    - [Adding Slide Numbers](#adding-slide-numbers)
    - [Slide Number Options](#slide-number-options)
    - [Slide Return Value](#slide-return-value)
  - [Saving a Presentation](#saving-a-presentation)
    - [Client Browser](#client-browser)
    - [Node.js](#nodejs-1)
    - [Saving Multiple Presentations](#saving-multiple-presentations)
- [Presentations: Adding Objects](#presentations-adding-objects)
  - [Adding Charts](#adding-charts)
    - [Chart Types](#chart-types)
    - [Multi-Type Charts](#multi-type-charts)
    - [Chart Size/Formatting Options](#chart-sizeformatting-options)
    - [Chart Axis Options](#chart-axis-options)
    - [Chart Data Options](#chart-data-options)
    - [Chart Element Shadow Options](#chart-element-shadow-options)
    - [Chart Multi-Type Options](#chart-multi-type-options)
    - [Chart Examples](#chart-examples)
  - [Adding Text](#adding-text)
    - [Text Options](#text-options)
    - [Text Shadow Options](#text-shadow-options)
    - [Text Examples](#text-examples)
  - [Adding Tables](#adding-tables)
    - [Table Layout Options](#table-layout-options)
    - [Table Auto-Paging Options](#table-auto-paging-options)
    - [Table Auto-Paging Notes](#table-auto-paging-notes)
    - [Table Formatting Options](#table-formatting-options)
    - [Table Formatting Notes](#table-formatting-notes)
    - [Table Cell Formatting](#table-cell-formatting)
    - [Table Cell Formatting Examples](#table-cell-formatting-examples)
    - [Table Examples](#table-examples)
  - [Adding Shapes](#adding-shapes)
    - [Shape Options](#shape-options)
    - [Shape Examples](#shape-examples)
  - [Adding Images](#adding-images)
    - [Image Options](#image-options)
    - [Image Examples](#image-examples)
    - [Image Sizing](#image-sizing)
  - [Adding Media (Audio/Video/YouTube)](#adding-media-audiovideoyoutube)
    - [Media Options](#media-options)
    - [Media Examples](#media-examples)
- [Master Slides and Corporate Branding](#master-slides-and-corporate-branding)
  - [Slide Masters](#slide-masters)
  - [Slide Master Object Options](#slide-master-object-options)
  - [Slide Master Examples](#slide-master-examples)
- [Table-to-Slides Feature](#table-to-slides-feature)
  - [Table-to-Slides Options](#table-to-slides-options)
  - [Table-to-Slides HTML Options](#table-to-slides-html-options)
  - [Table-to-Slides Notes](#table-to-slides-notes)
  - [Table-to-Slides Examples](#table-to-slides-examples)
  - [Creative Solutions](#creative-solutions)
- [Full PowerPoint Shape Library](#full-powerpoint-shape-library)
- [Scheme Colors](#scheme-colors)
- [Performance Considerations](#performance-considerations)
  - [Pre-Encode Large Images](#pre-encode-large-images)
- [Integration with Other Libraries](#integration-with-other-libraries)
  - [Integration with Angular](#integration-with-angular)
  - [Integration with Webpack/Typescript](#integration-with-webpacktypescript)
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
# Presentations: Adding Objects

Objects on the Slide are ordered from back-to-front based upon the order they were added.

For example, if you add an Image, then a Shape, then a Textbox: the Textbox will be in front of the Shape,
which is in front of the Image.


**************************************************************************************************



**************************************************************************************************


**************************************************************************************************


**************************************************************************************************
## Adding Shapes
Syntax (no text):
```javascript
slide.addShape({SHAPE}, {OPTIONS});
```
Syntax (with text):
```javascript
slide.addText("some string", {SHAPE, OPTIONS});
```
Check the `pptxgen.shapes.js` file for a complete list of the hundreds of PowerPoint shapes available.

### Shape Options
| Option       | Type    | Unit   | Default | Description         | Possible Values                                                             |
| :----------- | :------ | :----- | :------ | :------------------ | :-------------------------------------------------------------------------- |
| `x`          | number  | inches | `1.0`   | horizontal location | 0-n OR 'n%'. (Ex: `{x:'50%'}` will place object in the middle of the Slide) |
| `y`          | number  | inches | `1.0`   | vertical location   | 0-n OR 'n%'.                                                                |
| `w`          | number  | inches |         | width               | 0-n OR 'n%'. (Ex: `{w:'50%'}` will make object 50% width of the Slide)      |
| `h`          | number  | inches |         | height              | 0-n OR 'n%'.                                                                |
| `align`      | string  |        | `left`  | alignment           | `left` or `center` or `right`                                               |
| `fill`       | string  |        |         | fill/bkgd color     | hex color code or [scheme color constant](#scheme-colors). Ex: `{color:'0088CC'}` |
| `fill`       | object  |        |         | fill/bkgd color  | object with `type`, `color`, `alpha` (opt). Ex: `fill:{type:'solid', color:'0088CC', alpha:25}` |
| `flipH`      | boolean |        |         | flip Horizontal     | `true` or `false` |
| `flipV`      | boolean |        |         | flip Vertical       | `true` or `false` |
| `line`       | string  |        |         | border line color   | hex color code or [scheme color constant](#scheme-colors). Ex: `{line:'0088CC'}` |
| `lineDash`   | string  |        | `solid` | border line dash style | `dash`, `dashDot`, `lgDash`, `lgDashDot`, `lgDashDotDot`, `solid`, `sysDash` or `sysDot` |
| `line_head`   | string  |        |         | border line ending  | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none` |
| `lineSize`   | number  | points |         | border line size    | 1-256. Ex: {lineSize:4} |
| `line_tail`  | string  |        |         | border line heading | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none` |
| `rectRadius` | number  | inches |         | rounding radius     | rounding radius for `ROUNDED_RECTANGLE` text shapes |
| `rotate`     | integer | degrees|         | rotation degrees    | 0-360. Ex: `{rotate:180}` |

### Shape Examples
```javascript
var pptx = new PptxGenJS();
pptx.setLayout('LAYOUT_WIDE');

var slide = pptx.addNewSlide();

// LINE
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:4.40, w:5, h:0, line:'FF0000', lineSize:1 });
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:4.80, w:5, h:0, line:'FF0000', lineSize:2, line_head:'triangle' });
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:5.20, w:5, h:0, line:'FF0000', lineSize:3, line_tail:'triangle' });
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:5.60, w:5, h:0, line:'FF0000', lineSize:4, line_head:'triangle', line_tail:'triangle' });
// DIAGONAL LINE
slide.addShape(pptx.shapes.LINE,      { x:0, y:0, w:5.0, h:0, line:'FF0000', rotate:45 });
// RECTANGLE
slide.addShape(pptx.shapes.RECTANGLE, { x:0.50, y:0.75, w:5, h:3, fill:'FF0000' });
// OVAL
slide.addShape(pptx.shapes.OVAL,      { x:4.15, y:0.75, w:5, h:2, fill:{ type:'solid', color:'0088CC', alpha:25 } });

// Adding text to Shapes:
slide.addText('RIGHT-TRIANGLE', { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:0.40, y:4.3, w:6, h:3, fill:'0088CC', line:'000000', lineSize:3 });
slide.addText('RIGHT-TRIANGLE', { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:7.00, y:4.3, w:6, h:3, fill:'0088CC', line:'000000', flipH:true });

pptx.save('Demo-Shapes');
```


**************************************************************************************************
## Adding Images
Syntax:
```javascript
slide.addImage({OPTIONS});
```

Animated GIFs can be included in Presentations in one of two ways:
* Using Node.js: use either `data` or `path` options (Node can encode any image into base64)
* Client Browsers: pre-encode the gif and add it using the `data` option (encoding images into GIFs is beyond any current browser)

### Image Options
| Option       | Type    | Unit   | Default  | Description         | Possible Values  |
| :----------- | :------ | :----- | :------- | :------------------ | :--------------- |
| `x`          | number  | inches | `1.0`    | horizontal location | 0-n |
| `y`          | number  | inches | `1.0`    | vertical location   | 0-n |
| `w`          | number  | inches | `1.0`    | width               | 0-n |
| `h`          | number  | inches | `1.0`    | height              | 0-n |
| `data`       | string  |        |          | image data (base64) | base64-encoded image string. (either `data` or `path` is required) |
| `hyperlink`  | string  |        |          | add hyperlink | object with `url` or `slide` (`tooltip` optional). Ex: `{ hyperlink:{url:'https://github.com'} }` |
| `path`       | string  |        |          | image path          | Same as used in an (img src="") tag. (either `data` or `path` is required) |
| `sizing`     | object  |        |          | transforms image    | See [Image Sizing](#image-sizing) |

**NOTES**
* SVG images are not currently supported in PowerPoint or PowerPoint Online (even when encoded into base64). PptxGenJS does
properly encode and include SVG images, so they will begin showing once Microsoft adds support for this image type.
* Using `path` to add remote images (images from a different server) is not currently supported.

**Deprecation Warning**
Old positional parameters (e.g.: `slide.addImage('images/chart.png', 1, 1, 6, 3)`) are now deprecated as of 1.1.0

### Image Examples
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();

// Image by path
slide.addImage({ path:'images/chart_world_peace_near.png', x:1.0, y:1.0, w:8.0, h:4.0 });
// Image by data (base64-encoding)
slide.addImage({ data:'image/png;base64,iVtDafDrBF[...]=', x:3.0, y:5.0, w:6.0, h:3.0 });

// NOTE: Slide API calls return the same slide, so you can chain calls:
slide.addImage({ path:'images/cc_license_comp_chart.png', x:6.6, y:0.75, w:6.30, h:3.70 })
     .addImage({ path:'images/cc_logo.jpg',               x:0.5, y:3.50, w:5.00, h:3.70 })
     .addImage({ path:'images/cc_symbols_trans.png',      x:6.6, y:4.80, w:6.30, h:2.30 });

// Image with Hyperlink
slide.addImage({
  x:1.0, y:1.0, w:8.0, h:4.0,
  hyperlink:{ url:'https://github.com/gitbrent/pptxgenjs', tooltip:'Visit Homepage' },
  path:'images/chart_world_peace_near.png',
});

pptx.save('Demo-Images');
```

### Image Sizing
The `sizing` option provides cropping and scaling an image to a specified area. The property expects an object with the following structure:

| Property     | Type    | Unit   | Default           | Description                                   | Possible Values  |
| :----------- | :------ | :----- | :---------------- | :-------------------------------------------- | :--------------- |
| `type`       | string  |        |                   | sizing algorithm                              | `'crop'`, `'contain'` or `'cover'` |
| `w`          | number  | inches | `w` of the image  | area width                                    | 0-n |
| `h`          | number  | inches | `h` of the image  | area height                                   | 0-n |
| `x`          | number  | inches | `0`               | area horizontal position related to the image | 0-n (effective for `crop` only) |
| `y`          | number  | inches | `0`               | area vertical position related to the image   | 0-n (effective for `crop` only)|

Particular `type` values behave as follows:
* `contain` works as CSS property `background-size` — shrinks the image (ratio preserved) to the area given by `w` and `h` so that the image is completely visible. If the area's ratio differs from the image ratio, an empty space will surround the image.
* `cover` works as CSS property `background-size` — shrinks the image (ratio preserved) to the area given by `w` and `h` so that the area is completely filled. If the area's ratio differs from the image ratio, the image is centered to the area and cropped.
* `crop` cuts off a part specified by image-related coordinates `x`, `y` and size `w`, `h`.

NOTES:
* If you specify an area size larger than the image for the `contain` and `cover` type, then the image will be stretched, not shrunken.
* In case of the `crop` option, if the specified area reaches out of the image, then the covered empty space will be a part of the image.
* When the `sizing` property is used, its `w` and `h` values represent the effective image size. For example, in the following snippet, width and height of the image will both equal to 2 inches and its top-left corner will be located at [1 inch, 1 inch]:
```javascript
slide.addImage({
  path: '...',
  w: 4,
  h: 3,
  x: 1,
  y: 1,
  sizing: {
    type: 'contain',
    w: 2,
    h: 2
  }
});
```
**************************************************************************************************
## Adding Media (Audio/Video/YouTube)
Syntax:
```javascript
slide.addMedia({OPTIONS});
```

Both Video (mpg, mov, mp4, m4v, etc.) and Audio (mp3, wav, etc.) are supported (list of [supported formats](https://support.office.com/en-us/article/Video-and-audio-file-formats-supported-in-PowerPoint-d8b12450-26db-4c7b-a5c1-593d3418fb59#OperatingSystem=Windows))
* Using Node.js: use either `data` or `path` options (Node can encode any media into base64)
* Client Browsers: pre-encode the media and add it using the `data` option (encoding video/audio is beyond any current browser)

Online video (YouTube embeds, etc.) is supported in both client browser and in Node.js

### Media Options
| Option       | Type    | Unit   | Default   | Description         | Possible Values  |
| :----------- | :------ | :----- | :-------- | :------------------ | :--------------- |
| `x`          | number  | inches | `1.0`     | horizontal location | 0-n |
| `y`          | number  | inches | `1.0`     | vertical location   | 0-n |
| `w`          | number  | inches | `1.0`     | width               | 0-n |
| `h`          | number  | inches | `1.0`     | height              | 0-n |
| `data`       | string  |        |           | media data (base64) | base64-encoded string |
| `path`       | string  |        |           | media path          | relative path to media file |
| `link`       | string  |        |           | online url/link     | link to online video. Ex: `link:'https://www.youtube.com/embed/blahBlah'` |
| `type`       | string  |        |           | media type          | media type: `audio` or `video` (reqs: `data` or `path`) or `online` (reqs:`link`) |

### Media Examples
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();

// Media by path (Node.js only)
slide.addMedia({ type:'audio', path:'../media/sample.mp3', x:1.0, y:1.0, w:3.0, h:0.5 });
// Media by data (client browser or Node.js)
slide.addMedia({ type:'audio', data:'audio/mp3;base64,iVtDafDrBF[...]=', x:3.0, y:1.0, w:6.0, h:3.0 });
// Online by link (client browser or Node.js)
slide.addMedia({ type:'online', link:'https://www.youtube.com/embed/Dph6ynRVyUc', x:1.0, y:4.0, w:8.0, h:4.5 });

pptx.save('Demo-Media');
```


**************************************************************************************************
# Master Slides and Corporate Branding

## Slide Masters
Generating sample slides like those shown above is great for demonstrating library features,
but the reality is most of us will be required to produce presentations that have a certain design or
corporate branding.

PptxGenJS allows you to define Slide Master Layouts via objects that can then be used to provide branding
functionality.

Slide Masters are created by calling the `defineSlideMaster()` method along with an options object
(same style used in Slides).  Once defined, you can pass the Master title to `addNewSlide()` and that Slide will
use the Layout previously defined.  See the demo under /examples for several working examples.

The defined Masters become first-class Layouts in the exported PowerPoint presentation and can be changed
via View > Slide Master and will affect the Slides created using that layout.

## Slide Master Object Options
| Option        | Type    | Unit   | Default  | Description  | Possible Values       |
| :------------ | :------ | :----- | :------- | :----------- | :-------------------- |
| `bkgd`        | string  |        | `ffffff` | color        | hex color code or [scheme color constant](#scheme-colors). Ex: `{ bkgd:'0088CC' }` |
| `bkgd`        | object  |        |          | image | object with path OR data. Ex: `{path:'img/bkgd.png'}` OR `{data:'image/png;base64,iVBORwTwB[...]='}` |
| `slideNumber` | object  |        |          | Show slide numbers | ex: `{ x:1.0, y:'50%' }` `x` and `y` can be either inches or percent |
| `margin`      | number  | inches | `1.0`    | Slide margins      | 0.0 through Slide.width |
| `margin`      | array   |        |          | Slide margins      | array of numbers in TRBL order. Ex: `[0.5, 0.75, 0.5, 0.75]` |
| `objects`     | array   |        |          | Objects for Slide  | object with type and options. Type:`chart`,`image`,`line`,`rect` or `text`. [Example](https://github.com/gitbrent/PptxGenJS#slide-master-examples) |
| `title`       | string  |        |          | Layout title/name  | some title |

**TIP:**
Pre-encode your images (base64) and add the string as the optional data key/val (see `bkgd` above)

## Slide Master Examples
```javascript
var pptx = new PptxGenJS();
pptx.setLayout('LAYOUT_WIDE');

pptx.defineSlideMaster({
  title: 'MASTER_SLIDE',
  bkgd:  'FFFFFF',
  objects: [
    { 'line':  { x: 3.5, y:1.00, w:6.00, line:'0088CC', lineSize:5 } },
    { 'rect':  { x: 0.0, y:5.30, w:'100%', h:0.75, fill:'F1F1F1' } },
    { 'text':  { text:'Status Report', options:{ x:3.0, y:5.30, w:5.5, h:0.75 } } },
    { 'image': { x:11.3, y:6.40, w:1.67, h:0.75, path:'images/logo.png' } }
  ],
  slideNumber: { x:0.3, y:'90%' }
});

var slide = pptx.addNewSlide('MASTER_SLIDE');
slide.addText('How To Create PowerPoint Presentations with JavaScript', { x:0.5, y:0.7, font_size:18 });

pptx.save();
```




**************************************************************************************************
# Table-to-Slides Feature
Syntax:
```javascript
slide.addSlidesForTable(htmlElementID);
slide.addSlidesForTable(htmlElementID, {OPTIONS});
```

Any variety of HTML tables can be turned into a series of slides (auto-paging) by providing the table's ID.
* Reproduces an HTML table - background colors, borders, fonts, padding, etc.
* Slide margins are based on either the Master Slide provided or options

*NOTE: Nested tables are not supported in PowerPoint, so only the string contents of a single level deep table cell will be reproduced*

## Table-to-Slides Options
| Option            | Type    | Unit   | Description                     | Possible Values  |
| :---------------- | :------ | :----- | :------------------------------ | :--------------------------------------------- |
| `x`               | number  | inches | horizontal location             | 0-256. Table will be placed here on each Slide |
| `y`               | number  | inches | vertical location               | 0-256. Table will be placed here on each Slide |
| `w`               | number  | inches | width                           | 0-256. Default is (100% - Slide margins)       |
| `h`               | number  | inches | height                          | 0-256. Default is (100% - Slide margins)       |
| `master`          | string  |        | master slide to use             | [Slide Masters](#slide-masters) name. Ex: `{ master:'TITLE_SLIDE' }` |
| `addHeaderToEach` | boolean |        | add table headers to each slide | Ex: `addHeaderToEach:true`   |
| `addImage`        | string  |        | add an image to each slide      | Ex: `{ addImage:{ path:"images/logo.png", x:10, y:0.5, w:1.2, h:0.75 } }` |
| `addShape`        | string  |        | add a shape to each slide       | Use the established syntax   |
| `addTable`        | string  |        | add a table to each slide       | Use the established syntax   |
| `addText`         | string  |        | add text to each slide          | Use the established syntax   |

## Table-to-Slides HTML Options
Add an `data` attribute to the table's `<th>` tag to manually size columns (inches)
* minimum column width can be specified by using the `data-pptx-min-width` attribute
* fixed column width can be specified by using the `data-pptx-width` attribute

Example:
```HTML
<table id="tabAutoPaging" class="tabCool">
  <thead>
    <tr>
      <th data-pptx-min-width="0.6" style="width: 5%">Row</th>
      <th data-pptx-min-width="0.8" style="width:10%">Last Name</th>
      <th data-pptx-min-width="0.8" style="width:10%">First Name</th>
      <th data-pptx-width="8.5"     style="width:75%">Description</th>
    </tr>
  </thead>
  <tbody></tbody>
</table>
```

## Table-to-Slides Notes
* Default `x`, `y` and `margin` value is 0.5 inches, the table will take up all remaining space by default (h:100%, w:100%)
* Your Master Slides should already have defined margins, so a Master Slide name is the only option you'll need most of the time
* Hidden tables wont auto-size their columns correctly (as the properties are not accurate)

## Table-to-Slides Examples
```javascript
// Pass table element ID to addSlidesForTable function to produce 1-N slides
pptx.addSlidesForTable( 'myHtmlTableID' );

// Optionally, include a Master Slide name for pre-defined margins, background, logo, etc.
pptx.addSlidesForTable( 'myHtmlTableID', { master:'MASTER_SLIDE' } );

// Optionally, add images/shapes/text/tables to each Slide
pptx.addSlidesForTable( 'myHtmlTableID', { addText:{ text:"Dynamic Title", options:{x:1, y:0.5, color:'0088CC'} } } );
pptx.addSlidesForTable( 'myHtmlTableID', { addImage:{ path:"images/logo.png", x:10, y:0.5, w:1.2, h:0.75 } } );
```

## Creative Solutions
Design a Master Slide that already contains: slide layout, margins, logos, etc., then you can produce
professional looking Presentations with a single line of code which can be embedded into a link or a button:

Add a button to a webpage that will create a Presentation using whatever table data is present:
```html
<input type="button" value="Export to PPTX" onclick="{ var pptx=new PptxGenJS(); pptx.addSlidesForTable('tableId'); pptx.save(); }">
```

**SharePoint Integration**

Placing a button like this into a WebPart is a great way to add "Export to PowerPoint" functionality
to SharePoint. (You'd also need to add the 4 `<script>` includes in the same or another WebPart)

**************************************************************************************************
# Full PowerPoint Shape Library
If you are planning on creating Shapes (basically anything other than Text, Tables or Rectangles), then you'll want to
include the `pptxgen.shapes.js` library.

The shapes file contains a complete PowerPoint Shape object array thanks to the [officegen project](https://github.com/Ziv-Barber/officegen).

```javascript
<script lang="javascript" src="PptxGenJS/dist/pptxgen.shapes.js"></script>
```

**************************************************************************************************
# Scheme Colors
Scheme color is a variable that changes its value whenever another scheme palette is selected. Using scheme colors, design consistency can be easily preserved throughout the presentation and viewers can change color theme without any text/background contrast issues.

To use a scheme color, set a color constant as a property value:
```javascript
slide.addText('Hello',  { color: pptx.colors.TEXT1 });
```

The colors file contains a complete PowerPoint palette definition.

```javascript
<script lang="javascript" src="PptxGenJS/dist/pptxgen.colors.js"></script>
```

**************************************************************************************************
# Performance Considerations
It takes CPU time to read and encode images! The more images you include and the larger they are, the more time will be consumed.
The time needed to read/encode images can be completely eliminated by pre-encoding any images (see below).

## Pre-Encode Large Images
Pre-encode images into a base64 string (eg: 'image/png;base64,iVBORw[...]=') for use as the `data` option value.
This will both reduce dependencies (who needs another image asset to keep track of?) and provide a performance
boost (no time will need to be consumed reading and encoding the image).

**************************************************************************************************
# Integration with Other Libraries

## Integration with Angular

Set the browser mode option so the library will use blob file saving instead of detecting your app as a Node.js app (aka: avoid using `fs.writeFile`).  
* `pptx.setBrowser(true);`

[See Issue #220 for more information](https://github.com/gitbrent/PptxGenJS/issues/220)

## Integration with Webpack/Typescript

* Add to webpack config to avoid a module resolution error: `node: { fs: "empty" }`  
* Set browser mode so files will save as blobs via browser: `pptx.setBrowser(true);`

[See Issue #72 for more information](https://github.com/gitbrent/PptxGenJS/issues/72)

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
