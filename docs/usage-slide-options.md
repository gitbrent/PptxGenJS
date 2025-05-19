---
id: usage-slide-options
title: Slide Properties and Methods
---

## Slide Properties

| Option               | Type                                                                 | Default  | Description                      | Possible Values                                                                                                                                                    |
| :------------------- | :------------------------------------------------------------------- | :------- | :------------------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `background`         | `BackgroundProps`                                                    | `FFFFFF` | background color/images          | add background color or image [`DataOrPathProps`](./types.md#datapath-props-dataorpathprops) and/or [`ShapeFillProps`](./types.md#shape-fill-props-shapefillprops) |
| `color`              | string                                                               | `000000` | default text color               | hex color or [scheme color](./shapes-and-schemes.md).                                                                                                              |
| `hidden`             | boolean                                                              | `false`  | whether slide is hidden          | Ex: `slide.hidden = true`                                                                                                                                          |
| `newAutoPagedSlides` | PresSlide[]                                                          |          | all slides created by autopaging | Contains slides automatically created when content (e.g. a table) overflows the current slide using autoPage:true                                                  |
| `slideNumber`        | [`SlideNumberProps`](./types.md#slide-number-props-slidenumberprops) |          | slide number props               | (see examples below)                                                                                                                                               |

## Full Examples

### Example: Background/Foreground

```javascript
// EX: Use several methods to set a background
slide.background = { color: "F1F1F1" }; // Solid color
slide.background = { color: "FF3399", transparency: 50 }; // hex fill color with transparency of 50%
slide.background = { data: "image/png;base64,ABC[...]123" }; // image: base64 data
slide.background = { path: "https://some.url/image.jpg" }; // image: url
```

```javascript
// EX: Set slide default font color
slide.color = "696969";
```

### Example: Slide Number

```javascript
// EX: Add a Slide Number at a given location
slide.slideNumber = { x: 1.0, y: "90%" };

// EX: Styled Slide Numbers
slide.slideNumber = { x: 1.0, y: "95%", fontFace: "Courier", fontSize: 32, color: "CF0101" };
```
