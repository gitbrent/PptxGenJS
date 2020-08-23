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
slide.addText("some string", { SHAPE_NAME, SHAPE_PROPS });
```

## Shape Props

| Option         | Type           | Unit    | Default | Description          | Possible Values                                                         |
| :------------- | :------------- | :------ | :------ | :------------------- | :---------------------------------------------------------------------- |
| `x`            | number         | inches  | `1.0`   | horizontal location  | 0-n OR 'n%'. (Ex: `{x:'50%'}` places object in the middle of the Slide) |
| `y`            | number         | inches  | `1.0`   | vertical location    | 0-n OR 'n%'.                                                            |
| `w`            | number         | inches  |         | width                | 0-n OR 'n%'. (Ex: `{w:'50%'}` makes object 50% width of the Slide)      |
| `h`            | number         | inches  |         | height               | 0-n OR 'n%'.                                                            |
| `align`        | string         |         | `left`  | alignment            | `left` or `center` or `right`                                           |
| `fill`         | ShapeFillProps |         |         | fill props           | (see `ShapeFillProps`)                                                  |
| `flipH`        | boolean        |         |         | flip Horizontal      | `true` or `false`                                                       |
| `flipV`        | boolean        |         |         | flip Vertical        | `true` or `false`                                                       |
| `hyperlink`    | HyperlinkProps |         |         | hyperlink props      | (see `HyperlinkProps`)                                                  |
| `line`         | ShapeLineProps |         |         | border line props    | (see `ShapeLineProps`)                                                  |
| `rectRadius`   | number         | inches  |         | rounding radius      | 0-180. (only for pptx.shapes.ROUNDED_RECTANGLE)                         |
| `rotate`       | integer        | degrees | `0`     | rotation degrees     | -360 to 360. Ex: `{rotate:180}`                                         |
| `transparency` | number         |         | `0`     | transparency percent | 0-100                                                                   |
| `shadow`       | ShadowProps    |         |         | shadow props         | (see `ShadowProps`)                                                     |
| `shapeName`    | string         |         |         | name of shape        | optional name for shape, Ex: "Customer Network Diagram 99"              |

## Fill Props (`ShapeFillProps`)

| Option         | Type   | Default  | Description  | Possible Values                                                                  |
| :------------- | :----- | :------- | :----------- | :------------------------------------------------------------------------------- |
| `color`        | string | `000000` | color        | hex color code or [scheme color constant](#scheme-colors). Ex: `{line:'0088CC'}` |
| `transparency` | number | `0`      | transparency | Percentage: 0-100                                                                |

## Hyperlink Props (`HyperlinkProps`)

| Option    | Type   | Description           | Possible Values                      |
| :-------- | :----- | :-------------------- | :----------------------------------- |
| `slide`   | number | link to a given slide | Target Slide Number. Ex: `{slide:2}` |
| `tooltip` | string | link tooltip text     | Ex: `Click to visit home page`       |
| `url`     | string | target URL            | Ex: `https://wikipedia.org`          |

## Line Props (`ShapeLineProps`)

| Option           | Type   | Unit   | Default | Description       | Possible Values                                                                          |
| :--------------- | :----- | :----- | :------ | :---------------- | :--------------------------------------------------------------------------------------- |
| `beginArrowType` | string |        |         | line ending       | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none`                              |
| `color`          | string |        |         | line color        | hex color code or [scheme color constant](#scheme-colors). Ex: `{line:'0088CC'}`         |
| `dashType`       | string |        | `solid` | line dash style   | `dash`, `dashDot`, `lgDash`, `lgDashDot`, `lgDashDotDot`, `solid`, `sysDash` or `sysDot` |
| `endArrowType`   | string |        |         | line heading      | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none`                              |
| `transparency`   | number |        | `0`     | line transparency | Percentage: 0-100                                                                        |
| `width`          | number | points | `1`     | line width/size   | 1-256. Ex: `{ width:4 }`                                                                 |

## Shadow Props (`ShadowProps`)

| Option    | Type   | Default  | Description            | Possible Values          |
| :-------- | :----- | :------- | :--------------------- | :----------------------- |
| `type`    | string | `none`   | shadow type            | `outer`, `inner`, `none` |
| `angle`   | number | `0`      | blue degrees           | `0`-`359`                |
| `blur`    | number | `0`      | blur range (points)    | `0`-`100`                |
| `color`   | string | `000000` | color                  | hex color code           |
| `offset`  | number | `0`      | shadow offset (points) | `0`-`200`                |
| `opacity` | number | `0`      | opacity percentage     | `0.0`-`1.0`              |

## Shape Examples

```javascript
// Plain shapes:
let slide1 = pres.addSlide();
slide1.addShape(pres.ShapeType.rect, { x: 0.5, y: 0.8, w: 1.5, h: 3.0, fill: { color: "FF0000" } });
slide1.addShape(pres.ShapeType.ellipse, { x: 5.4, y: 0.8, w: 3.0, h: 1.5, fill: { type: "solid", color: "0088CC" } });
slide1.addShape(pres.ShapeType.line, { x: 4.2, y: 4.4, w: 5.0, h: 0.0, line: { color: "FF0000", width: 1 } });
slide1.addShape(pres.ShapeType.line, { x: 4.2, y: 4.8, w: 5.0, h: 0.0, line: { color: "FF0000", width: 2, beginArrowType: "triangle" } });

// Shapes with Text:
let slide2 = pres.addSlide();
slide2.addText("RECTANGLE", { shape: pres.ShapeType.rect, x: 0.5, y: 0.8, w: 1.5, h: 3.0, fill: { color: "FF0000" }, align: "center", fontSize: 14 });
slide2.addText("ELLIPSE", { shape: pres.ShapeType.ellipse, x: 5.4, y: 0.8, w: 3.0, h: 1.5, fill: { color: "FF0000" }, align: "center", fontSize: 14 });
slide2.addText("LINE size=1", { shape: pres.ShapeType.line, align: "center", x: 4.2, y: 4.4, w: 5, h: 0, line: { color: "FF0000", width: 1, dashType: "lgDash" } });
slide2.addText("LINE size=2", { shape: pres.ShapeType.line, align: "left", x: 4.2, y: 4.8, w: 5, h: 0, line: { color: "FF0000", width: 2, endArrowType: "triangle" } });

pres.writeFile("Demo-Shapes");
```
