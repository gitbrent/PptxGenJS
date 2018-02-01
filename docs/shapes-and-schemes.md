---
id: shapes-and-schemes
title: Shapes and Schemes
---
**************************************************************************************************
Table of Contents
- [PowerPoint Shapes](#powerpoint-shape-library)
- [PowerPoint Scheme Colors](#powerpoint-scheme-colors)
**************************************************************************************************

## PowerPoint Shapes
If you are planning on creating Shapes (basically anything other than Text, Tables or Rectangles), then you'll want to
include the `pptxgen.shapes.js` library.

The shapes file contains a complete PowerPoint Shape object array thanks to the [officegen project](https://github.com/Ziv-Barber/officegen).

```javascript
<script lang="javascript" src="PptxGenJS/dist/pptxgen.shapes.js"></script>
```

## PowerPoint Scheme Colors
Scheme color is a variable that changes its value whenever another scheme palette is selected. Using scheme colors, design consistency can be easily preserved throughout the presentation and viewers can change color theme without any text/background contrast issues.

To use a scheme color, set a color constant as a property value:
```javascript
slide.addText('Hello', { color:pptx.colors.TEXT1 });
```

The colors file contains a complete PowerPoint palette definition.

```javascript
<script lang="javascript" src="PptxGenJS/dist/pptxgen.colors.js"></script>
```
