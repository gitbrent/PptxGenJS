# PptxGenJS
A complete JavaScript PowerPoint Presentation creator framework for client web browsers.

By including the PptxGenJS framework inside an HTML page, you have the ability to quickly and
easily produce PowerPoint presentations with a few simple JavaScript commands.
* Works with all modern desktop browsers (IE11, Edge, Chrome, Firefox, Opera)
* The presentation export is pushed to client browsers as a regular file without any interaction required
* Complete HTML5/JavaScript solution - no other libraries, plug-ins, or settings are required

Additionally, this framework includes a unique Table-To-Slides feature that will reproduce
an HTML Table in 1 or more slides with a single command.

Supported write/output formats:
* PowerPoint 2007+ and Open Office XML format (.PPTX)

# Demo
[http://gitbrent.github.io/PptxGenJS](http://gitbrent.github.io/PptxGenJS)

# Installation
PptxGenJS requires just three other libraries to produce and push a file to web browsers.

```javascript
<script lang="javascript" src="dist/jquery.min.js"></script>
<script lang="javascript" src="dist/filesaver.js"></script>
<script lang="javascript" src="dist/jszip.min.js"></script>
<script lang="javascript" src="dist/pptxgen.min.js"></script>
```

# Optional Modules
If you are planning on creating Shapes (basically anything other than Text, Tables or rectangles), then you'll want to
include the `pptxgen.shapes.js` library.  It's a complete PowerPoint PPTX Shape object array thanks to the
[officegen project](https://github.com/Ziv-Barber/officegen)
```javascript
<script lang="javascript" src="dist/pptxgen.shapes.js"></script>
```

For times when a plain white slide won't due, why not create a Slide Master? (especially useful for corporate environments).
See the one in the examples folder to get started
```javascript
<script lang="javascript" src="dist/pptxgen.masters.js"></script>
```

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->
**Table of Contents**  *generated with [DocToc](https://github.com/thlorenz/doctoc)*

- [PptxGenJS](#pptxgenjs)
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
- [Library Reference](#library-reference)
  - [Presentation Options](#presentation-options)
  - [Available Layouts](#available-layouts)
  - [Creating Slides](#creating-slides)
  - [Text](#text)
    - [Text Options](#text-options)
  - [Table](#table)
    - [Table Options](#table-options)
  - [Shape](#shape)
  - [Image](#image)
  - [Performance Considerations](#performance-considerations)
- [Bugs & Issues](#bugs-&-issues)
- [License](#license)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

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
## TIP:
* Placing a button like this into a WebPart is a great way to add "Export to PowerPoint" functionality
to SharePoint/Office365. (You'd also need to add the 4 `<script>` includes in the same or another WebPart)

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

slide.addImage('images/cc_copyremix.gif',          0.5, 0.75, 2.35, 2.45 );
// Slide API calls return the same slide, so you can chain calls:
slide.addImage('images/cc_license_comp_chart.png', 6.6, 0.75, 6.30, 3.70 )
     .addImage('images/cc_logo.jpg',               0.5, 3.50, 5.00, 3.70 )
     .addImage('images/cc_symbols_trans.png',      6.6, 4.80, 6.30, 2.30 );

pptx.save('Demo-Shapes');
```

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
| LAYOUT_WIDE  | No      | 13.3 x 7.5 inches |
| LAYOUT_4x3   | No      | 10 x 7.5 inches   |
| LAYOUT_16x10 | No      | 10 x 6.25 inches  |
| LAYOUT_16x9  | Yes     | 10 x 5.625 inches |

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
slide.addText('World!', { x:2.7, y:1.0, color:'DDDD00', font_size:90 });
```

### Text Options
| Parameter  | Description    | Possible Values       |
| :--------- | :------------- | :-------------------- |
| x          | X location     | (inches)              |
| y          | Y location     | (inches)              |
| inset      | inset/padding  | (inches)              |
| align      | horiz align    | left / center / right |
| valign     | vert align     | top / middle / bottom |
| autoFit    | "Fit to Shape" | true / false          |

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

## Shape
```javascript
// Syntax
slide.addShape({SHAPE}, {options});

// Example: Red Rectangle
slide.addShape(pptx.shapes.RECTANGLE, { x:0.50, y:0.75, cx:5, cy:3.2, fill:'FF0000' });
// View the pptxgen.shapes.js file for a complete list of Shapes
```

## Image
```javascript
// Syntax
slide.addShape({SHAPE}, {options});

// Example: Located at 1.5" x 1.5", 6.0" Wide, 3.0" Height
slide.addImage('images/cc_license_comp_chart.png', 1.5, 1.5, 6.0, 3.0);
```

## Performance Considerations
NOTE: It takes time to encode images, so the more images you include and the larger they are, the more time will be consumed.
You will want to show a jQuery Dialog with a nice hour glass before you start creating the file.



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
