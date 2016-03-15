# PptxGenJS
A complete JavaScript PowerPoint Presentation creator framework for client web browsers.

By including the PptxGenJS framework inside an HTML page, you have the ability to quickly and easily produce PowerPoint presentations with a few simple JavaScript commands.
* Works with all modern desktop browsers (IE11, Edge, Chrome, Firefox, Opera)
* The presentation export is pushed to client browsers as a regular file without any interaction required
* Complete HTML5/JavaScript solution - no other libraries, plug-ins, or settings are required

Supported write/output formats:
* PowerPoint 2007+ and Open Office XML format (.PPTX)

# Demo
[http://gitbrent.github.io/PptxGenJS](http://gitbrent.github.io/PptxGenJS)

# Installation
* PptxGenJS requires just a few libraries to produce and push a file to web browsers
```javascript
<script lang="javascript" src="dist/filesaver.js"></script>
<script lang="javascript" src="dist/jquery.min.js"></script>
<script lang="javascript" src="dist/jszip.min.js"></script>
<script lang="javascript" src="dist/pptxgen.min.js"></script>
```

# Optional Modules
* A complete PowerPoint Shape object array (thanks to )
```javascript
<script lang="javascript" src="dist/pptxgen.masters.js"></script>
<script lang="javascript" src="dist/pptxgen.shapes.js"></script>
```

# Writing Presentation
Creating a Presentation is super easy.

### Creating your first Presentation:

```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();
slide.addText('Hello World!', { x:0.5, y:0.7, font_size:18, font_face:'Arial', color:'0000FF' });
pptx.save('HelloWorld');
```

See, that was easy!

### The Basics:
* Options are passed via objects
* Height, width and locations are passed in inches
* Not much other than X and Y locations are required

### Presentation Options

Setting the Title:
```javascript
pptx.setPresTitle('PptxGenJS Sample Export');
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

### Slides
* Presentations are composed of 1 or more Slides

Creating a new Slide
```javascript
var slide = pptx.addNewSlide();
```

(*Optional*) Slides can take a single argument: the name of a Master Slide to use.
```javascript
var slide = pptx.addNewSlide(TITLE_SLIDE);
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

## Lastly
NOTE: It takes time to encode images, so the more images you include and the larger they are, the more time will be consumed.
You will want to show a jQuery Dialog with a nice hour glass before you start creating the file.



# Bugs & Issues

When reporting bugs or issues, if you could include a link to a simple jsbin or similar demonstrating the issue, that'd be really helpful.

# License

[MIT License](http://opensource.org/licenses/MIT)

Copyright (c) 2010-2016 Brent Ely, [https://github.com/GitBrent/PptxGenJS](https://github.com/GitBrent/PptxGenJS)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
