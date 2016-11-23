# PptxGenJS
Client-side JavaScript framework that produces PowerPoint (pptx) presentations.

By including the PptxGenJS framework inside an HTML page, you have the ability to quickly and
easily produce PowerPoint presentations with a few simple JavaScript commands.
* Works with all modern desktop browsers (IE11, Edge, Chrome, Firefox, Opera)
* The presentation export is pushed to client browsers as a regular file without any interaction required
* Complete, modern JavaScript solution - no client configuration, plug-ins, or other settings needed!

Additionally, this framework includes a unique Table-To-Slides feature that will reproduce
an HTML Table in 1 or more slides with a single command.

Supported write/output formats:
* PowerPoint 2007+, Open Office XML, Apple Keynote (.PPTX)

Now available on NPM/Node:
* [https://www.npmjs.com/package/pptxgenjs](https://www.npmjs.com/package/pptxgenjs)

**************************************************************************************************

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->
**Table of Contents**  (*generated with [DocToc](https://github.com/thlorenz/doctoc)*)

- [Demo](#demo)
- [Installation](#installation)
- [Optional Modules](#optional-modules)
- [The Basics](#the-basics)
- [Creating a Presentation](#creating-a-presentation)
- [Table-to-Slides / 1-Click Exports](#table-to-slides--1-click-exports)
  - [TIP:](#tip)
- [In-Depth Examples](#in-depth-examples)
  - [Table Example](#table-example)
  - [Text Example](#text-example)
  - [Shape Example](#shape-example)
  - [Image Example](#image-example)
- [Master Slides and Corporate Branding](#master-slides-and-corporate-branding)
  - [Slide Masters](#slide-masters)
  - [Slide Master Examples](#slide-master-examples)
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
  - [Performance Considerations](#performance-considerations)
- [Bugs & Issues](#bugs-&-issues)
- [License](#license)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

**************************************************************************************************

# Demo
Use JavaScript to Create PowerPoint presentations right from our demo page  
[http://gitbrent.github.io/PptxGenJS](http://gitbrent.github.io/PptxGenJS)

# Installation
PptxGenJS requires only three additional JavaScript libraries to function.

```javascript
<script lang="javascript" src="PptxGenJS/libs/jquery.min.js"></script>
<script lang="javascript" src="PptxGenJS/libs/jszip.min.js"></script>
<script lang="javascript" src="PptxGenJS/libs/filesaver.min.js"></script>
<script lang="javascript" src="PptxGenJS/dist/pptxgen.js"></script>
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

## TIP:
* Placing a button like this into a WebPart is a great way to add "Export to PowerPoint" functionality
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
It's one thing to generate sample slides like those shown above, but most of us are required to produce
slides that have a consistent look and feel and/or with corporate branding, etc.

In addition, it's often useful to have various pre-defined slide masters to use as giant "Thank You"
slides, title slides, etc.

Fortunately, you can do both with PptxGenJS!

Slide Masters are defined using the same style as the Slides and then added as a variable to a file that
is included in the script src tags on your page.  
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
    images:     [ { src:'images/logo_square.png', x:9.3, y:4.9, w:0.5, h:0.5 } ],
    shapes:     [
      { type:'text', text:'ACME - Confidential', x:0, y:5.17, cx:'100%', cy:0.3, align:'center', valign:'top', color:'7F7F7F', font_size:8, bold:true },
      { type:'line', x:0.3, y:3.85, cx:5.7, cy:0.0, line:'007AAA' },
      { type:'rectangle', x:0, y:0, w:'100%', h:.65, cx:5, cy:3.2, fill:'003b75' }
    ]
  },
  TITLE_SLIDE: {
    title:      'I am the Title Slide',
    isNumbered: false,
    bkgd:       { src:'images/title_bkgd.png', data:'base64,R0lGONlhotPQBMAPyoAPosR[...]+0pEZbEhAAOw==' },
    images:     [ { x:'7.4', y:'4.1', w:'2', h:'1', data:'data:image/png;base64,R0lGODlhPQBEAPeoAJosM[...]+0pCZbEhAAOw==' } ]
  }
};
```  
#### PRO-TIP: Pre-encode Images for Performance Boost
Pre-encode your images (base64) and add the string as the optional data key/val
(see the `TITLE_SLIDE.images` object above)

Every object added to the global master slide variable `gObjPptxMasters` can then be referenced
by their key names that you created (e.g.: "TITLE_SLIDE").  

```javascript
var pptx = new PptxGenJS();

var slide1 = pptx.addNewSlide( pptx.masters.TITLE_SLIDE );
slide1.addText('How To Create PowerPoint Presentations with JavaScript', { x:0.5, y:0.7, font_size:18 });
// NOTE: Base master slide properties can be overridden on a selective basis:
// Here we can set a new background color or image on-the-fly
var slide2 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:'0088CC'} );
var slide3 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:{ src:'images/title_bkgd.jpg' } } );

pptx.save();
```

## Slide Master Object Options
| Parameter  | Description        | Possible Values       |
| :--------- | :-------------     | :-------------------- |
| bkgd       | background color   | [string] CSS-Hex style ('336699', etc.) OR [object] {src:'img/path'} - (optional) data:'base64code' |
| images     | image(s)           | array of image objects: {src,x,y,cx,cy} - (optional) data:'base64code' |
| isNumbered | Show slide numbers | true/false |
| margin     | slide margin       | [integer] OR [array] of ints in TRBL order [top,right,bottom,left] (inches) |
| shapes     | shape(s)           | array of shape objects: {type,text,x,y,cx,cy,align,valign,color,font_size,bold} |
| title      | Slide Title        | [text string] |

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
| Layout Name  | Default | Description       |
| :----------- | :-------| :---------------- |
| LAYOUT_16x9  | Yes     | 10 x 5.625 inches |
| LAYOUT_16x10 | No      | 10 x 6.25 inches  |
| LAYOUT_4x3   | No      | 10 x 7.5 inches   |
| LAYOUT_WIDE  | No      | 13.3 x 7.5 inches |

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
| Parameter  | Description    | Possible Values       |
| :--------- | :------------- | :-------------------- |
| x          | X location     | (inches)              |
| y          | Y location     | (inches)              |
| w          | width          | (inches)              |
| h          | height         | (inches)              |
| align      | horiz align    | left / center / right |
| autoFit    | "Fit to Shape" | true / false          |
| bold       | bold           | true/false            |
| bullet     | bullet text    | true/false            |
| color      | font color     | CSS-Hex style ('336699', etc.) |
| fill       | fill color     | CSS-Hex style ('336699', etc.) |
| font_face  | font face      | 'Arial' etc. |
| font_size  | font size      | Std PPT font sizes (1-256 pt) |
| inset      | inset/padding  | (inches)              |
| isTextBox  | PPT "Textbox"  | true/false |
| italic     | italic         | true/false |
| margin     | margin/padding | Points/Pixels (use same value as CSS passing for similar results in PPT) |
| underline  | underline      | true/false |
| valign     | vertical align | top / middle / bottom |

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
| Parameter  | Description    | Possible Values       |
| :--------- | :------------- | :-------------------- |
| x          | X location     | (inches)              |
| y          | Y location     | (inches)              |
| w          | table width    | (inches)              |
| h          | table height   | (inches)              |

### Cell Options
| Parameter  | Description    | Possible Values       |
| :--------- | :------------- | :-------------------- |
| align      | horiz align    | left / center / right |
| bold       | bold           | true/false |
| border     | cell border    | single style or an array (TRBL-format) of styles - {pt:'1', color:'CFCFCF'} |
| color      | font color     | CSS-Hex style ('336699', etc.) |
| colspan    | column span    | 2-N |
| colW       | column width   | an int or array of ints in TRBL order [top,right,bottom,left] (inches) |
| font_face  | font face      | 'Arial' etc. |
| font_size  | font size      | Std PPT font sizes (1-256 pt) |
| fill       | fill color     | CSS-Hex style ('336699', etc.) |
| italic     | italic         | true/false |
| marginPt   | margin/padding | Points/Pixels (use same value as CSS passing for similar results in PPT) |
| rowspan    | row span       | 2-N |
| rowH       | row height     | an int or array of ints in TRBL order [top,right,bottom,left] (inches) |
| underline  | underline      | true/false |
| valign     | vert align     | top / middle / bottom |


## Shape
```javascript
// Syntax
slide.addShape({SHAPE}, {options});

// Example: Red Rectangle
slide.addShape(pptx.shapes.RECTANGLE, { x:0.50, y:0.75, cx:5, cy:3.2, fill:'FF0000' });
// View the pptxgen.shapes.js file for a complete list of Shapes
```

### Shape Options
| Parameter  | Description     | Possible Values       |
| :--------- | :-------------- | :-------------------- |
| x          | X location      | (inches)              |
| y          | Y location      | (inches)              |
| w          | shape width     | (inches)              |
| h          | shape height    | (inches)              |
| flipH      | flip Horizontal | true/false            |
| flipV      | flip Vertical   | true/false            |
| rotate     | rotate          | 0-360 (integer)       |

## Image
```javascript
// Syntax
slide.addImage({options});

// Example: Image by path / Image by base64-encoding
slide.addImage({ path:'images/chart_world_peace_is_close.png', x:1.0, y:1.0, w:8.0, h:4.0 });
slide.addImage({ data:'data:image/png;base64,iVBORwTwB[...]=', x:3.0, y:5.0, w:6.0, h:3.0 });
```

### Image Options
| Parameter  | Description    | Possible Values       |
| :--------- | :------------- | :-------------------- |
| path       | image path     | (path - can be relative - like a normal html tag: img src="path")
| data       | image data     | (base64-encoded string) |
| x          | X location     | (inches)              |
| y          | Y location     | (inches)              |
| w          | image width    | (inches)              |
| h          | image height   | (inches)              |

### Deprecation Warning
Old positional parameters (e.g.: `slide.addImage('images/chart.png', 1.5, 1.5, 6.0, 3.0)`) are now deprecated

**************************************************************************************************
## Performance Considerations
NOTE: It takes time to encode images, so the more images you include and the larger they are, the more time will be consumed.
You will want to show a jQuery Dialog with a nice hour glass before you start creating the file.

## PRO-TIP
Pre-encode images into a base64 string (eg: 'data:image/png;base64,iVBORw[...]=') and add as the data
argument so exports are super fast (no need to read/encode images!) and reduce dependencies (you dont
need yet another img asset to keep track of or deploy right?)

**************************************************************************************************
# Bugs & Issues

When reporting bugs or issues, if you could include a link to a simple jsbin or similar demonstrating the issue, that'd be really helpful.

**************************************************************************************************
# License

[MIT License](http://opensource.org/licenses/MIT)

Copyright (c) 2015-2016 Brent Ely, [https://github.com/GitBrent/PptxGenJS](https://github.com/GitBrent/PptxGenJS)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
