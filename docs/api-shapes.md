---
id: api-shapes
title: Shapes
---

Almost 200 shape types can be added to Slides (see [`ShapeType`](https://github.com/gitbrent/PptxGenJS/blob/master/types/index.d.ts) enum).

## Usage

```typescript
// Shapes without text
slide.addShape(pres.ShapeType.rect, { fill: { color: "FF0000" } });
slide.addShape(pres.ShapeType.ellipse, {
  fill: { type: "solid", color: "0088CC" },
});
slide.addShape(pres.ShapeType.line, { line: { color: "FF0000", width: 1 } });

// Shapes with text
slide.addText("ShapeType.rect", {
  shape: pres.ShapeType.rect,
  fill: { color: "FF0000" },
});
slide.addText("ShapeType.ellipse", {
  shape: pres.ShapeType.ellipse,
  fill: { color: "FF0000" },
});
slide.addText("ShapeType.line", {
  shape: pres.ShapeType.line,
  line: { color: "FF0000", width: 1, dashType: "lgDash" },
});
```

## Properties

### Position/Size Props ([PositionProps](/PptxGenJS/docs/types#position-props))

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

### Shape Props ([ShapeProps](/PptxGenJS/docs/types#shape-props-shapeprops))

| Name         | Type                                                                    | Description         | Possible Values                                             |
| :----------- | :---------------------------------------------------------------------- | :------------------ | :---------------------------------------------------------- |
| `align`      | string                                                                  | alignment           | `left` or `center` or `right`. Default: `left`              |
| `fill`       | [ShapeFillProps](/PptxGenJS/docs/types#shape-fill-props-shapefillprops) | fill props          | Solid or gradient fill props                                |
| `flipH`      | boolean                                                                 | flip Horizontal     | `true` or `false`                                           |
| `flipV`      | boolean                                                                 | flip Vertical       | `true` or `false`                                           |
| `hyperlink`  | [HyperlinkProps](/PptxGenJS/docs/types#hyperlink-props-hyperlinkprops)  | hyperlink props     | (see type link)                                             |
| `line`       | [ShapeLineProps](/PptxGenJS/docs/types#shape-line-props-shapelineprops) | border line props   | (see type link)                                             |
| `rectRadius` | number                                                                  | rounding radius     | 0 to 1. (Ex: 0.5. Only for `pptx.shapes.ROUNDED_RECTANGLE`) |
| `rotate`     | number                                                                  | rotation (degrees)  | -360 to 360. Default: `0`                                   |
| `shadow`     | [ShadowProps](/PptxGenJS/docs/types#shadow-props-shadowprops)           | shadow props        | (see type link)                                             |
| `shapeName`  | string                                                                  | optional shape name | Ex: "Customer Network Diagram 99"                           |

## Gradient Fills

Set `fill.type` to `"gradient"` and provide a [`ShapeGradientProps`](/PptxGenJS/docs/types#shape-gradient-props-shapegradientprops) object in `fill.gradient`. A gradient requires at least two color stops. If fewer are provided, PptxGenJS uses a solid fill instead.

### Linear Gradient

Linear gradients are the default. `angle` is measured clockwise in degrees: `0` is left-to-right and `90` is top-to-bottom. Each stop can also specify its own transparency percentage.

```typescript
slide.addShape(pres.ShapeType.rect, {
  x: 1,
  y: 1,
  w: 4,
  h: 2,
  fill: {
    type: "gradient",
    gradient: {
      angle: 45,
      stops: [
        { pos: 0, color: "4472C4" },
        { pos: 50, color: "70AD47" },
        { pos: 100, color: "FFC000", transparency: 25 },
      ],
    },
  },
});
```

### Radial Gradient

Set the gradient `type` to `"radial"`:

```typescript
slide.addShape(pres.ShapeType.ellipse, {
  x: 1,
  y: 3.5,
  w: 3,
  h: 2,
  fill: {
    type: "gradient",
    gradient: {
      type: "radial",
      stops: [
        { pos: 0, color: "FFFFFF" },
        { pos: 100, color: "4472C4" },
      ],
    },
  },
});
```

### Gradient Lines and Slide Backgrounds

The same gradient definition can be used with `line` or `slide.background`:

```typescript
slide.addShape(pres.ShapeType.line, {
  x: 5,
  y: 1,
  w: 4,
  h: 0,
  line: {
    type: "gradient",
    width: 4,
    gradient: {
      stops: [
        { pos: 0, color: "4472C4" },
        { pos: 100, color: "ED7D31" },
      ],
    },
  },
});

slide.background = {
  type: "gradient",
  gradient: {
    angle: 90,
    stops: [
      { pos: 0, color: "D9EAF7" },
      { pos: 100, color: "FFFFFF" },
    ],
  },
};
```

## Examples

![Shapes with Text Demo](./assets/ex-shape-slide.png)

## Samples

Sample code all available types: [demos/modules/demo_shape.mjs](https://github.com/gitbrent/PptxGenJS/blob/master/demos/modules/demo_shape.mjs)
