[![Open Source Love](https://badges.frapsoft.com/os/v1/open-source.svg?v=103)](https://github.com/ellerbrock/open-source-badge/) [![MIT Licence](https://badges.frapsoft.com/os/mit/mit.svg?v=103)](https://opensource.org/licenses/mit-license.php) [![npm version](https://badge.fury.io/js/pptxgenjs.svg)](https://badge.fury.io/js/pptxgenjs)
# PptxGenJS
Client-side JavaScript framework that produces PowerPoint (pptx) presentations.

Include the PptxGenJS framework inside an HTML page (or a node project), to gain the ability to quickly and
easily produce PowerPoint presentations with a few simple JavaScript commands.

## Main Features
* Complete, modern JavaScript solution - no client configuration, plug-ins, or other settings required
* Works with all modern desktop browsers (Chrome, Edge, Firefox, IE11, Opera et al.)
* Presentation pptx export is pushed to client browsers as a regular file without the need for any user interaction
* Simple, feature-rich API: Supports Master Slides and all major object types (Tables, Shapes, Images and Text)

## Additional Features
This framework also includes a unique [Table-to-Slides](#table-to-slides--1-click-exports) feature that will reproduce
an HTML Table across one or more Slides with a single command.

**************************************************************************************************

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->
**Table of Contents**  (*generated with [DocToc](https://github.com/thlorenz/doctoc)*)

- [Demo](#demo)
- [Installation](#installation)
  - [Client-Side](#client-side)
  - [NPM/Node.js](#npmnodejs)
- [Optional Modules](#optional-modules)
- [The Basics](#the-basics)
- [Creating a Presentation](#creating-a-presentation)
- [Table-to-Slides / 1-Click Exports](#table-to-slides--1-click-exports)
  - [Slide Branding](#slide-branding)
    - [ProTip](#protip)
- [In-Depth Examples](#in-depth-examples)
  - [Table Example](#table-example)
  - [Text Example](#text-example)
  - [Shape Example](#shape-example)
  - [Image Example](#image-example)
- [Master Slides and Corporate Branding](#master-slides-and-corporate-branding)
  - [Slide Masters](#slide-masters)
  - [Slide Master Examples](#slide-master-examples)
    - [ProTip](#protip-1)
  - [Slide Master Object Options](#slide-master-object-options)
  - [Sample Slide Master File](#sample-slide-master-file)
- [Library Reference](#library-reference)
  - [Presentation Options](#presentation-options)
  - [Available Layouts](#available-layouts)
  - [Creating Slides](#creating-slides)
  - [Text](#text)
    - [Text Options](#text-options)
  - [Table](#table)
    - [Table Options](#table-options)
    - [Cell Options](#cell-options)
  - [Shape](#shape)
    - [Shape Options](#shape-options)
  - [Image](#image)
    - [Image Options](#image-options)
    - [Deprecation Warning](#deprecation-warning)
- [Performance Considerations](#performance-considerations)
  - [Pre-Encode Large Images](#pre-encode-large-images)
- [Bugs & Issues](#bugs-&-issues)
- [Special Thanks](#special-thanks)
- [License](#license)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

**************************************************************************************************

# Demo
Use JavaScript to Create PowerPoint presentations right from our demo page  
[http://gitbrent.github.io/PptxGenJS](http://gitbrent.github.io/PptxGenJS)

# Installation
## Client-Side
PptxGenJS requires only three additional JavaScript libraries to function.
```javascript
<script lang="javascript" src="PptxGenJS/libs/jquery.min.js"></script>
<script lang="javascript" src="PptxGenJS/libs/jszip.min.js"></script>
<script lang="javascript" src="PptxGenJS/libs/filesaver.min.js"></script>
<script lang="javascript" src="PptxGenJS/dist/pptxgen.js"></script>
```
## NPM/Node.js
```javascript
npm install pptxgenjs
```

# Optional Modules
If you are planning on creating Shapes (basically anything other than Text, Tables or Rectangles), then you'll want to
include the `pptxgen.shapes.js` library.  It's a complete PowerPoint PPTX Shape object array thanks to the
[officegen project](https://github.com/Ziv-Barber/officegen)
```javascript
<script lang="javascript" src="PptxGenJS/dist/pptxgen.shapes.js"></script>
```

**************************************************************************************************
# The Basics
* Presentations are composed of 1 or more Slides
* Options are passed via objects (e.g.: `{ x:1.5, y:2.5, font_size:18 }`)
* Shape dimensions and locations are passed in inches
* Not much other than X and Y locations are required are required for most objects

# Creating a Presentation
Creating a Presentation is as easy as 1-2-3:

1. Add a Slide  
2. Add any Shapes, Text or Tables  
3. Save the Presentation  

```javascript
var slide = pptx.addNewSlide();
slide.addText('Hello World!', { x:0.5, y:0.7, font_size:18, color:'0000FF' });
pptx.save('Sample Presentation');
```

**************************************************************************************************
# Table-to-Slides / 1-Click Exports
* With the unique `addSlidesForTable()` function, you can reproduce an HTML table - background
colors, borders, fonts, padding, etc. - with a single function call.
* The function will detect margins (based on Master Slide layout or parameters) and will create Slides as needed
* All you have to do is throw a table at the function and you're done!

```javascript
// STEP 1: Instantiate new PptxGenJS instance
var pptx = new PptxGenJS();

// STEP 2: Set slide size/layout
pptx.setLayout('LAYOUT_16x9');

// STEP 3: Pass table to addSlidesForTable function to produce 1-N slides
pptx.addSlidesForTable('tabAutoPaging');

// STEP 4: Export Presentation
pptx.save('Table2SlidesDemo');
```

## Slide Branding
What about cases where you have a specific Slide Master or Corporate layout to adhere to?  
No problem!  
Simply pass the Slide Master name and all shapes/text will appear on the output Slides.  Even better,
your slide layout/size and margins are already defined as well, so you end up with code you can just inline
into a button and place next to any table on your site.

```javascript
<input type="button" value="Export to PPTX" onclick="{ var pptx = new PptxGenJS(); pptx.addSlidesForTable('tableId',{ master:pptx.masters.MASTER_SLIDE }); pptx.save(); }">
```

Note: Slide background color/image can be overridden on a per-slide basis when needed.
```javascript
var slide1 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:'0088CC'} );
```

### ProTip
Placing a button like this into a WebPart is a great way to add "Export to PowerPoint" functionality
to SharePoint. (You'd also need to add the 4 `<script>` includes in the same or another WebPart)

**************************************************************************************************
# In-Depth Examples

## Table Example
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();
slide.addText('Demo-03: Table', { x:0.5, y:0.25, font_size:18, font_face:'Arial', color:'0088CC' });

// TABLE 1: Simple array
// --------
var rows = [ 1,2,3,4,5,6,7,8,9,10 ];
var tabOpts = { x:0.5, y:1.0, cx:9.0 };
var celOpts = { fill:'F7F7F7', font_size:14, color:'363636' };
slide.addTable( rows, tabOpts, celOpts );

// TABLE 2: Multi-row Array
// --------
var rows = [
    ['A1', 'B1', 'C1'],
    ['A2', 'B2', 'C3']
];
var tabOpts = { x:0.5, y:2.0, cx:9.0 };
var celOpts = { fill:'dfefff', font_size:18, color:'6f9fc9', rowH:1.0, valign:'m', align:'c', border:{pt:'1', color:'FFFFFF'} };
slide.addTable( rows, tabOpts, celOpts );

// TABLE 3: Formatting on a cell-by-cell basis - (TIP: use this to over-ride table options)
// --------
var rows = [
    [
        { text: 'Top Lft', opts: { valign:'t', align:'l', font_face:'Arial'   } },
        { text: 'Top Ctr', opts: { valign:'t', align:'c', font_face:'Verdana' } },
        { text: 'Top Rgt', opts: { valign:'t', align:'r', font_face:'Courier' } }
    ],
];
var tabOpts = { x:0.5, y:4.5, cx:9.0 };
var celOpts = { fill:'dfefff', font_size:18, color:'6f9fc9', rowH:0.6, valign:'m', align:'c', border:{pt:'1', color:'FFFFFF'} };
slide.addTable( rows, tabOpts, celOpts );

pptx.save('Demo-Tables');
```

## Text Example
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();

slide.addText('Hello',  { x:0.5, y:0.7, cx:3, color:'0000FF', font_size:64 });
slide.addText('World!', { x:2.7, y:1.0, cx:5, color:'DDDD00', font_size:90 });
slide.addText('^ (50%/50%)', {x:'50%', y:'50%'});
var objOptions = {
    x:0.5, y:4.3, cx:'90%',
    font_face:'Arial', font_size:32, color:'00CC00', bold:true, underline:true, margin:0, isTextBox:true
};
slide.addText('Arial, 32pt, green, bold, underline, 0 inset', objOptions);

pptx.save('Demo-Text');
```

## Shape Example
```javascript
var pptx = new PptxGenJS();
pptx.setLayout('LAYOUT_WIDE');

var slide = pptx.addNewSlide();
// Misc Shapes
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:4.40, cx:5, cy:0, line:'FF0000', line_size:1 });
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:4.80, cx:5, cy:0, line:'FF0000', line_size:2, line_head:'triangle' });
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:5.20, cx:5, cy:0, line:'FF0000', line_size:3, line_tail:'triangle' });
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:5.60, cx:5, cy:0, line:'FF0000', line_size:4, line_head:'triangle', line_tail:'triangle' });
slide.addShape(pptx.shapes.OVAL,      { x:4.15, y:0.75, cx:5, cy:2.0, fill:{ type:'solid', color:'0088CC', alpha:25 } });
slide.addShape(pptx.shapes.RECTANGLE, { x:0.50, y:0.75, cx:5, cy:3.2, fill:'FF0000' });
// Add text to Shapes:
slide.addText('RIGHT-TRIANGLE', { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:0.40, y:4.3, cx:6, cy:3, fill:'0088CC', line:'000000', line_size:3 });
slide.addText('RIGHT-TRIANGLE', { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:7.00, y:4.3, cx:6, cy:3, fill:'0088CC', line:'000000', flipH:true });

pptx.save('Demo-Shapes');
```

## Image Example
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();

slide.addImage({ path:'images/cc_copyremix.gif',          x:0.5, y:0.75, w:2.35, h:2.45 });
// Slide API calls return the same slide, so you can chain calls:
slide.addImage({ path:'images/cc_license_comp_chart.png', x:6.6, y:0.75, w:6.30, h:3.70 })
     .addImage({ path:'images/cc_logo.jpg',               x:0.5, y:3.50, w:5.00, h:3.70 })
     .addImage({ path:'images/cc_symbols_trans.png',      x:6.6, y:4.80, w:6.30, h:2.30 });

pptx.save('Demo-Shapes');
```

**************************************************************************************************
# Master Slides and Corporate Branding

## Slide Masters
Generating sample slides like those shown above is great for demonstrating library features,
but the reality is most of us will be required to produce presentations that have a certain design or
corporate branding.

PptxGenJS allows you to define Master Slides via objects that can then be used to provide branding
functionality.

Slide Masters are defined using the same object style used in Slides. Add these objects as a variable to a file that
is included in the script src tags on your page, then reference them by name in your code.  
E.g.: `<script lang="javascript" src="pptxgenjs.masters.js"></script>`

## Slide Master Examples
`pptxgenjs.masters.js` contents:
```javascript
var gObjPptxMasters = {
  MASTER_SLIDE: {
    title:      'Slide master',
    isNumbered: true,
    margin:     [ 0.5, 0.25, 1.0, 0.25 ],
    bkgd:       'FFFFFF',
    images:     [ { path:'images/logo_square.png', x:9.3, y:4.9, w:0.5, h:0.5 } ],
    shapes:     [
      { type:'text', text:'ACME - Confidential', x:0, y:5.17, cx:'100%', cy:0.3, align:'center', valign:'top', color:'7F7F7F', font_size:8, bold:true },
      { type:'line', x:0.3, y:3.85, cx:5.7, cy:0.0, line:'007AAA' },
      { type:'rectangle', x:0, y:0, w:'100%', h:.65, cx:5, cy:3.2, fill:'003b75' }
    ]
  },
  TITLE_SLIDE: {
    title:      'I am the Title Slide',
    isNumbered: false,
    bkgd:       { data:'image/png;base64,R0lGONlhotPQBMAPyoAPosR[...]+0pEZbEhAAOw==' },
    images:     [ { x:'7.4', y:'4.1', w:'2', h:'1', data:'image/png;base64,R0lGODlhPQBEAPeoAJosM[...]+0pCZbEhAAOw==' } ]
  }
};
```  
Every object added to the global master slide variable `gObjPptxMasters` can then be referenced
by their key names that you created (e.g.: "TITLE_SLIDE").  

### ProTip
Pre-encode your images (base64) and add the string as the optional data key/val
(see the `TITLE_SLIDE.images` object above)

```javascript
var pptx = new PptxGenJS();

var slide1 = pptx.addNewSlide( pptx.masters.TITLE_SLIDE );
slide1.addText('How To Create PowerPoint Presentations with JavaScript', { x:0.5, y:0.7, font_size:18 });
// NOTE: Base master slide properties can be overridden on a selective basis:
// Here we can set a new background color or image on-the-fly
var slide2 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:'0088CC' } );
var slide3 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:{ path:'images/title_bkgd.jpg' } } );
var slide4 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:{ data:'image/png;base64,tFfInmP[...]=' } } );

pptx.save();
```

## Slide Master Object Options
| Option       | Type    | Unit   | Default  | Description  | Possible Values       |
| :----------- | :------ | :----- | :------- | :----------- | :-------------------- |
| `bkgd`       | string  |        | `ffffff` | color        | hex color code. Ex: `{ bkgd:'0088CC' }` |
| `bkgd`       | object  |        |          | image | object with path OR data. Ex: `{path:'img/bkgd.png'}` OR `{data:'image/png;base64,iVBORwTwB[...]='}` |
| `images`     | array   |        |          | image(s) | object array of path OR data. Ex: `{path:'img/logo.png'}` OR `{data:'image/png;base64,tFfInmP[...]'}`|
| `isNumbered` | boolean |        | `false`  | Show slide numbers | `true` or `false` |
| `margin`     | number  | inches | `1.0`    | Slide margin       | 0.0 through whatever |
| `margin`     | array   |        |          | Slide margins      | array of numbers in TRBL order. Ex: `[0.5, 0.75, 0.5, 0.75]` |
| `shapes`     | array   |        |          | shape(s)           | array of shape objects. Ex: (see [Shape](#shape) section) |
| `title`      | string  |        |          | Slide title        | some title |

## Sample Slide Master File
A sample masters file is included in the distribution folder and contain a couple of different slides to get you started.  
Location: `PptxGenJS/dist/pptxgen.masters.js`

**************************************************************************************************
# Library Reference

## Presentation Options
Setting the Title:
```javascript
pptx.setTitle('PptxGenJS Sample Export');
```
Setting the Layout (layout is applied to every Slide in the Presentation):
```javascript
pptx.setLayout('LAYOUT_WIDE');
```

## Available Layouts
| Layout Name    | Default  | Layout Slide Size |
| :------------- | :------- | :---------------- |
| `LAYOUT_16x9`  | Yes      | 10 x 5.625 inches |
| `LAYOUT_16x10` | No       | 10 x 6.25 inches  |
| `LAYOUT_4x3`   | No       | 10 x 7.5 inches   |
| `LAYOUT_WIDE`  | No       | 13.3 x 7.5 inches |

## Creating Slides

```javascript
var slide = pptx.addNewSlide();
```

(*Optional*) Slides can take a single argument: the name of a Master Slide to use.
```javascript
var slide = pptx.addNewSlide(pptx.masters.TITLE_SLIDE);
```

## Text
```javascript
// Syntax
slide.addText('TEXT', {OPTIONS});

// Example
slide.addText('World!', { x:2.5, y:3.5, color:'DDDD00', font_size:90 });
slide.addText('Options!', { x:1, y:1, font_face:'Arial', font_size:42, color:'00CC00', bold:true, italic:true, underline:true } );
```

### Text Options
| Option       | Type    | Unit   | Default   | Description | Possible Values  |
| :----------- | :------ | :----- | :-------- | :---------- | :--------------- |
| `x`          | number  | inches | `1.0`     | horizontal location | 0-n |
| `y`          | number  | inches | `1.0`     | vertical location   | 0-n |
| `w`          | number  | inches |           | width               | 0-n |
| `h`          | number  | inches |           | height              | 0-n |
| `align`      | string  |        | `left`    | alignment       | `left` or `center` or `right` |
| `autoFit`    | boolean |        | `false`   | "Fit to Shape"  | `true` or `false` |
| `bold`       | boolean |        | `false`   | bold text       | `true` or `false` |
| `bullet`     | boolean |        | `false`   | bullet text     | `true` or `false` |
| `color`      | string  |        |           | text color      | hex color code. Ex: `{ color:'0088CC' }` |
| `fill`       | string  |        |           | fill/bkgd color | hex color code. Ex: `{ color:'0088CC' }` |
| `font_face`  | string  |        |           | font face       | Ex: 'Arial' |
| `font_size`  | number  | points |           | font size       | 1-256. Ex: `{ font_size:12 }` |
| `inset`      | number  | inches | `1.0`     | inset/padding   | 1-256. Ex: `{ inset:10 }` |
| `isTextBox`  | boolean |        | `false`   | PPT "Textbox"   | `true` or `false` |
| `italic`     | boolean |        | `false`   | italic text     | `true` or `false` |
| `margin`     | number  | points |           | margin          | 1-n (ProTip: use the same value from CSS padding) |
| `underline`  | boolean |        | `false`   | underline text  | `true` or `false` |
| `valign`     | string  |        | `left`    | vertical alignment | `top` or `middle` or `bottom` |

## Table
```javascript
// Syntax
slide.addTable( [rows] );
slide.addTable( [rows], {tabOpts} );
slide.addTable( [rows], {tabOpts}, {cellOpts} );

// Basic Example
slide.addTable( ['A1', 'B1', 'C1'] );

// Cell-Styling Example
var rows = [
    [
        { text: 'Top Lft', opts: { valign:'t', align:'l', font_face:'Arial'   } },
        { text: 'Top Ctr', opts: { valign:'t', align:'c', font_face:'Verdana' } },
        { text: 'Top Rgt', opts: { valign:'t', align:'r', font_face:'Courier' } }
    ],
];
var cellOpts = { fill:'dfefff', font_size:18, color:'6f9fc9', rowH:0.6, valign:'m', align:'c', border:{pt:'1', color:'FFFFFF'} };
// The cellOpts provide a way to format all cells
// Individual cell opts override this base style, so you can quickly format a table with minimum effort
slide.addTable( rows, { x:0.5, y:4.5, cx:9.0 }, cellOpts );
```

### Table Options
| Option       | Type    | Unit   | Default   | Description         | Possible Values  |
| :----------- | :------ | :----- | :-------- | :------------------ | :--------------- |
| `x`          | number  | inches | `1.0`     | horizontal location | 0-n |
| `y`          | number  | inches | `1.0`     | vertical location   | 0-n |
| `w`          | number  | inches |           | width               | 0-n |
| `h`          | number  | inches |           | height              | 0-n |
| `colW`       | integer | inches |           | width for every column | Ex: Width for every column in table (uniform) `2.0` |
| `colW`       | array   | inches |           | column widths in order | Ex: Width for each of 5 columns `[1.0, 2.0, 2.5, 1.5, 1.0]` |
| `rowH`       | integer | inches |           | height for every row   | Ex: Height for every row in table (uniform) `2.0` |
| `rowH`       | array   | inches |           | row heights in order   | Ex: Height for each of 5 rows `[1.0, 2.0, 2.5, 1.5, 1.0]` |

### Cell Options
| Option       | Type    | Unit   | Default   | Description        | Possible Values  |
| :----------- | :------ | :----- | :-------- | :----------------- | :--------------- |
| `align`      | string  |        | `left`    | alignment          | `left` or `center` or `right` |
| `bold`       | boolean |        | `false`   | bold text          | `true` or `false` |
| `border`     | object  |        |           | cell border        | object with `pt` and `color` values. Ex: `{pt:'1', color:'CFCFCF'}` |
| `border`     | array   |        |           | cell border        | array of objects with `pt` and `color` values in TRBL order. |
| `color`      | string  |        |           | text color         | hex color code. Ex: `{color:'0088CC'}` |
| `colspan`    | integer |        |           | column span        | 2-n. Ex: `{colspan:2}` |
| `fill`       | string  |        |           | fill/bkgd color    | hex color code. Ex: `{color:'0088CC'}` |
| `font_face`  | string  |        |           | font face          | Ex: 'Arial' |
| `font_size`  | number  | points |           | font size          | 1-256. Ex: `{ font_size:12 }` |
| `italic`     | boolean |        | `false`   | italic text        | `true` or `false` |
| `marginPt`   | number  | points |           | margin             | 1-n (ProTip: use the same value from CSS padding) |
| `rowspan`    | integer |        |           | row span           | 2-n. Ex: `{rowspan:2}` |
| `underline`  | boolean |        | `false`   | underline text     | `true` or `false` |
| `valign`     | string  |        | `left`    | vertical alignment | `top` or `middle` or `bottom` |

## Shape
```javascript
// Syntax
slide.addShape({SHAPE}, {options});

// Example: Red Rectangle
slide.addShape(pptx.shapes.RECTANGLE, { x:0.50, y:0.75, cx:5, cy:3.2, fill:'FF0000' });
// View the pptxgen.shapes.js file for a complete list of Shapes
```

### Shape Options
| Option       | Type    | Unit   | Default   | Description         | Possible Values  |
| :----------- | :------ | :----- | :-------- | :------------------ | :--------------- |
| `x`          | number  | inches | `1.0`     | horizontal location | 0-n |
| `y`          | number  | inches | `1.0`     | vertical location   | 0-n |
| `w`          | number  | inches | `1.0`     | width               | 0-n |
| `h`          | number  | inches | `1.0`     | height              | 0-n |
| `flipH`      | boolean |        |           | flip Horizontal     | `true` or `false` |
| `flipV`      | boolean |        |           | flip Vertical       | `true` or `false` |
| `rotate`     | integer | degrees |          | rotation degrees    | 0-360. Ex: `{rotate:180}` |

## Image
```javascript
// Syntax
slide.addImage({options});

// Example: Image by path / Image by base64-encoding
slide.addImage({ path:'images/chart_world_peace_near.png', x:1.0, y:1.0, w:8.0, h:4.0 });
slide.addImage({ data:'image/png;base64,iVtDafDrBF[...]=', x:3.0, y:5.0, w:6.0, h:3.0 });
```

### Image Options
| Option       | Type    | Unit   | Default   | Description         | Possible Values  |
| :----------- | :------ | :----- | :-------- | :------------------ | :--------------- |
| `x`          | number  | inches | `1.0`     | horizontal location | 0-n |
| `y`          | number  | inches | `1.0`     | vertical location   | 0-n |
| `w`          | number  | inches | `1.0`     | width               | 0-n |
| `h`          | number  | inches | `1.0`     | height              | 0-n |
| `data`       | string  |        |           | image data (base64) | base64-encoded image string. (either `data` or `path` is required) |
| `path`       | string  |        |           | image path          | Same as used in an (img src="") tag. (either `data` or `path` is required) |

### Deprecation Warning
Old positional parameters (e.g.: `slide.addImage('images/chart.png', 1, 1, 6, 3)`) are now deprecated as of 1.1.0

**************************************************************************************************
# Performance Considerations
It takes time to read and encode images! The more images you include and the larger they are, the more time will be consumed.
You will want to show a jQuery Dialog with a nice hour glass before you start creating the file.

## Pre-Encode Large Images
Pre-encode images into a base64 string (eg: 'data:image/png;base64,iVBORw[...]=') and use as the `data` argument.
This will both reduce dependencies (who needs another image asset to keep track of?) and provide a performance
boost (no time will need to be consumed reading and encoding the image).

**************************************************************************************************
# Bugs & Issues

When reporting bugs or issues, if you could include a link to a simple jsbin or similar demonstrating the issue, that'd be really helpful.

**************************************************************************************************
# Special Thanks

* [Officegen Project](https://github.com/Ziv-Barber/officegen) - For the Shape definitions and XML code
* [Dzmitry Dulko](https://github.com/DzmitryDulko) - For getting the project published on NPM
* Everyone who has submitted a Patch or an Issue

**************************************************************************************************
# License

Copyright &copy; 2015-2016 [Brent Ely](https://github.com/gitbrent/PptxGenJS)

[MIT](https://github.com/gitbrent/PptxGenJS/blob/master/LICENSE)
