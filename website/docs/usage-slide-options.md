---
id: usage-slide-options
title: Slide Properties
---

## Slide Properties

| Option       | Type            | Default  | Description             | Possible Values                                                       |
| :----------- | :-------------- | :------- | :---------------------- | :-------------------------------------------------------------------- |
| `background` | BackgroundProps | `FFFFFF` | background color/images | add background color or image                                         |
| `color`      | string          | `000000` | default text color      | hex color or [scheme color](/PptxGenJS/docs/shapes-and-schemes.html). |
| `hidden`     | boolean         | `false`  | whether slide is hidden | Ex: `slide.hidden = true`                                             |

### Examples

```javascript
slide.background = { fill: "F1F1F1" }; // Solid color
slide.background = { data: "image/png;base64,ABC[...]123" }; // image: base64 data
slide.background = { path: "https://some.url/image.jpg" }; // image: url

slide.color = "696969"; // Set slide default font color
```

## Slide Numbers

### Syntax

```javascript
slide.slideNumber = { x: 1.0, y: "90%" };
```

## Properties

### Position Props ([PositionProps](/PptxGenJS/docs/types.html#position-props))

| Option | Type   | Default | Description            | Possible Values                              |
| :----- | :----- | :------ | :--------------------- | :------------------------------------------- |
| `x`    | number | `1.0`   | hor location (inches)  | 0-n                                          |
| `x`    | string |         | hor location (percent) | 'n%'. (Ex: `{x:'50%'}` middle of the Slide)  |
| `y`    | number | `1.0`   | ver location (inches)  | 0-n                                          |
| `y`    | string |         | ver location (percent) | 'n%'. (Ex: `{y:'50%'}` middle of the Slide)  |
| `w`    | number | `1.0`   | width (inches)         | 0-n                                          |
| `w`    | string |         | width (percent)        | 'n%'. (Ex: `{w:'50%'}` 50% the Slide width)  |
| `h`    | number | `1.0`   | height (inches)        | 0-n                                          |
| `h`    | string |         | height (percent)       | 'n%'. (Ex: `{h:'50%'}` 50% the Slide height) |

### Slide Number Props (`SlideNumberProps`)

| Option     | Type   | Default  | Description | Possible Values                                                       |
| :--------- | :----- | :------- | :---------- | :-------------------------------------------------------------------- |
| `color`    | string | `000000` | color       | hex color or [scheme color](/PptxGenJS/docs/shapes-and-schemes.html). |
| `fontFace` | string |          | font face   | any available font. Ex: `{ fontFace:'Arial' }`                        |
| `fontSize` | number |          | font size   | 8-256. Ex: `{ fontSize:12 }`                                          |

### Slide Number Examples

```javascript
// EX: Add a Slide Number at a given location
slide.slideNumber = { x: 1.0, y: "90%" };

// EX: Styled Slide Numbers
slide.slideNumber = { x: 1.0, y: "95%", fontFace: "Courier", fontSize: 32, color: "CF0101" };
```
