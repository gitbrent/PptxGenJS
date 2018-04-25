---
id: usage-add-slide
title: Adding a Slide
---

## Syntax
```javascript
var slide = pptx.addNewSlide();
```

## Slide Formatting
```javascript
slide.back  = 'F1F1F1';
slide.color = '696969';
```

## Slide Formatting Options
| Option       | Type    | Unit   | Default   | Description         | Possible Values  |
| :----------- | :------ | :----- | :-------- | :------------------ | :--------------- |
| `bkgd`       | string  |        | `FFFFFF`  | background color    | hex color code or [scheme color constant](#scheme-colors). |
| `color`      | string  |        | `000000`  | default text color  | hex color code or [scheme color constant](#scheme-colors). |

## Applying Master Slides / Branding
```javascript
// Create a new Slide that will inherit properties from a pre-defined master page (margins, logos, text, background, etc.)
var slide1 = pptx.addNewSlide('TITLE_SLIDE');

// The background color can be overridden on a per-slide basis:
var slide2 = pptx.addNewSlide('TITLE_SLIDE', {bkgd:'FFFCCC'});
```

## Adding Slide Numbers
```javascript
// EX: Basic Slide Numbers
slide.slideNumber({ x:1.0, y:'90%' });

// EX: Custom styled Slide Numbers
slide.slideNumber({ x:1.0, y:'90%', fontFace:'Courier', fontSize:32, color:'CF0101' });
```

## Slide Number Options
| Option       | Type    | Unit   | Default   | Description         | Possible Values  |
| :----------- | :------ | :----- | :-------- | :------------------ | :--------------- |
| `x`          | number  | inches | `0.3`     | horizontal location | 0-n OR 'n%'. (Ex: `{x:'10%'}` places number 10% from left edge) |
| `y`          | number  | inches | `90%`     | vertical location   | 0-n OR 'n%'. (Ex: `{y:'90%'}` places number 90% down the Slide) |
| `color`      | string  |        |           | text color          | hex color code or [scheme color constant](#scheme-colors). Ex: `{color:'0088CC'}` |
| `fontFace`   | string  |        |           | font face           | any available font. Ex: `{fontFace:Arial}` |
| `fontSize`   | number  | points |           | font size           | 8-256. Ex: `{fontSize:12}` |

## Slide Return Value
The Slide object returns a reference to itself, so calls can be chained.

Example:
```javascript
slide
.addImage({ path:'images/logo1.png', x:1, y:2, w:3, h:3 })
.addImage({ path:'images/logo2.jpg', x:5, y:3, w:3, h:3 })
.addImage({ path:'images/logo3.png', x:9, y:4, w:3, h:3 });
```
