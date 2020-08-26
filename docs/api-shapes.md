---
id: api-shapes
title: Adding Shapes
---

## Syntax

### Shape without text

```javascript
slide.addShape(SHAPE_NAME, SHAPE_PROPS);
```

### Shape with text

```javascript
slide.addText("This is a Triangle", { SHAPE_NAME, SHAPE_PROPS });
```

## Properties

### Position Props ([PositionProps](/PptxGenJS/docs/types.html#position-props))

| Name | Type   | Default | Description            | Possible Values                              |
| :--- | :----- | :------ | :--------------------- | :------------------------------------------- |
| `x`  | number | `1.0`   | hor location (inches)  | 0-n                                          |
| `x`  | string |         | hor location (percent) | 'n%'. (Ex: `{x:'50%'}` middle of the Slide)  |
| `y`  | number | `1.0`   | ver location (inches)  | 0-n                                          |
| `y`  | string |         | ver location (percent) | 'n%'. (Ex: `{y:'50%'}` middle of the Slide)  |
| `w`  | number | `1.0`   | width (inches)         | 0-n                                          |
| `w`  | string |         | width (percent)        | 'n%'. (Ex: `{w:'50%'}` 50% the Slide width)  |
| `h`  | number | `1.0`   | height (inches)        | 0-n                                          |
| `h`  | string |         | height (percent)       | 'n%'. (Ex: `{h:'50%'}` 50% the Slide height) |

### Shape Props (`ShapeProps`)

| Name         | Type                                                                         | Description         | Possible Values                                   |
| :----------- | :--------------------------------------------------------------------------- | :------------------ | :------------------------------------------------ |
| `align`      | string                                                                       | alignment           | `left` or `center` or `right`. Default: `left`    |
| `fill`       | [ShapeFillProps](/PptxGenJS/docs/types.html#fill-props-shapefillprops)       | fill props          | Fill color/transparency props                     |
| `flipH`      | boolean                                                                      | flip Horizontal     | `true` or `false`                                 |
| `flipV`      | boolean                                                                      | flip Vertical       | `true` or `false`                                 |
| `hyperlink`  | [HyperlinkProps](/PptxGenJS/docs/types.html#hyperlink-props-hyperlinkprops)  | hyperlink props     | (see type link)                                   |
| `line`       | [ShapeLineProps](/PptxGenJS/docs/types.html#shape-line-props-shapelineprops) | border line props   | (see type link)                                   |
| `rectRadius` | number                                                                       | rounding radius     | 0-180. (only for `pptx.shapes.ROUNDED_RECTANGLE`) |
| `rotate`     | number                                                                       | rotation (degrees)  | -360 to 360. Default: `0`                         |
| `shadow`     | [ShadowProps](/PptxGenJS/docs/types.html#shadow-props-shadowprops)           | shadow props        | (see type link)                                   |
| `shapeName`  | string                                                                       | optional shape name | Ex: "Customer Network Diagram 99"                 |

## Examples

### Shapes without text

```javascript
let slide1 = pres.addSlide();
slide1.addShape(pres.ShapeType.rect, { x: 0.5, y: 0.8, w: 1.5, h: 3.0, fill: { color: "FF0000" } });
slide1.addShape(pres.ShapeType.ellipse, { x: 5.4, y: 0.8, w: 3.0, h: 1.5, fill: { type: "solid", color: "0088CC" } });
slide1.addShape(pres.ShapeType.line, { x: 4.2, y: 4.4, w: 5.0, h: 0.0, line: { color: "FF0000", width: 1 } });
slide1.addShape(pres.ShapeType.line, { x: 4.2, y: 4.8, w: 5.0, h: 0.0, line: { color: "FF0000", width: 2, beginArrowType: "triangle" } });

pres.writeFile("Demo-Shapes-1");
```

### Shapes with text

```javascript
let slide2 = pres.addSlide();
slide2.addText("RECTANGLE", { shape: pres.ShapeType.rect, x: 0.5, y: 0.8, w: 1.5, h: 3.0, fill: { color: "FF0000" }, align: "center", fontSize: 14 });
slide2.addText("ELLIPSE", { shape: pres.ShapeType.ellipse, x: 5.4, y: 0.8, w: 3.0, h: 1.5, fill: { color: "FF0000" }, align: "center", fontSize: 14 });
slide2.addText("LINE size=1", { shape: pres.ShapeType.line, align: "center", x: 4.2, y: 4.4, w: 5, h: 0, line: { color: "FF0000", width: 1, dashType: "lgDash" } });
slide2.addText("LINE size=2", { shape: pres.ShapeType.line, align: "left", x: 4.2, y: 4.8, w: 5, h: 0, line: { color: "FF0000", width: 2, endArrowType: "triangle" } });

pres.writeFile("Demo-Shapes-2");
```
