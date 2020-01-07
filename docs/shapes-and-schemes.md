---
id: shapes-and-schemes
title: Shapes and Schemes
---

## PowerPoint Shapes

The library comes with over 180 built-in PowerPoint shapes.

- See [core-shapes.ts](https://github.com/gitbrent/PptxGenJS/blob/master/src/core-shapes.ts) for the complete list
- Compliments of the [officegen project](https://github.com/Ziv-Barber/officegen)

## PowerPoint Scheme Colors

Scheme color is a variable that changes its value whenever another scheme palette is selected. Using scheme colors, design consistency can be easily preserved throughout the presentation and viewers can change color theme without any text/background contrast issues.

To use a scheme color, set a color constant as a property value:

```javascript
slide.addText("Hello", { color: pptx.colors.TEXT1 });
```

See the [Shapes Demo](https://gitbrent.github.io/PptxGenJS/demo/#shapes) for Scheme Colors demo

![Scheme Demo](/PptxGenJS/docs/assets/demo-scheme.png)
