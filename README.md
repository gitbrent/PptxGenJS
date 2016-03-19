# PptxGenJS
A complete JavaScript PowerPoint Presentation creator framework for client web browsers.

By including the PptxGenJS framework inside an HTML page, you have the ability to quickly and easily produce PowerPoint presentations with a few simple JavaScript commands.
* Works with all modern desktop browsers (IE11, Edge, Chrome, Firefox, Opera)
* The presentation export is pushed to client browsers as a regular file without any interaction required
* Complete HTML5/JavaScript solution - no other libraries, plug-ins, or settings are required

Additionally, this framework includes a unique Table-To-Slides feature that will reproduce an HTML Table in 1 or more slides with a single command.

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

For times when a plain white slide won't due, why not create a Slide Master? (especially useful for corporate environments). See the one in the examples folder to get started
```javascript
<script lang="javascript" src="dist/pptxgen.masters.js"></script>
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
* With the unique `addSlidesForTable()` function, you can reproduce an HTML table - background colors, borders, fonts, padding, etc. - with a single function call.
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
### TIP:
* Placing a button like this into a WebPart is a great way to add "Export to PowerPoint" functionality to SharePoint/Office365. (You'd also need to add the 4 `<script>` includes in the same or another WebPart)

**************************************************************************************************
# In-Depth Examples

**************************************************************************************************
# Library Reference

### Presentation Options
Setting the Title:
```javascript
pptx.setTitle('PptxGenJS Sample Export');
```
Setting the Layout: (Layout applied to every Slide in the Presentation)
```javascript
pptx.setLayout('LAYOUT_WIDE');
```

### Available Layouts
| Layout Name  | Default | Description       |
| :----------- | :-------| :---------------- |
| LAYOUT_WIDE  | No      | 13.3 x 7.5 inches |
| LAYOUT_4x3   | No      | 10 x 7.5 inches   |
| LAYOUT_16x10 | No      | 10 x 6.25 inches  |
| LAYOUT_16x9  | Yes     | 10 x 5.625 inches |

### Creating Slides

```javascript
var slide = pptx.addNewSlide();
```

(*Optional*) Slides can take a single argument: the name of a Master Slide to use.
```javascript
var slide = pptx.addNewSlide(pptx.masters.TITLE_SLIDE);
```

## Text
```javascript
```

## Table
```javascript
```

## Shape
```javascript
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
