---
id: usage-slide-options
title: Slide Options
---

## Slide Numbers

### Slide Number Syntax
```javascript
slide.slideNumber({ x:1.0, y:'90%' });
```

### Slide Number Options
| Option       | Type    | Unit   | Default | Description         | Possible Values  |
| :----------- | :------ | :----- | :------ | :------------------ | :-------------------------------------------------------------- |
| `x`          | number  | inches | `0.3`   | horizontal location | 0-n OR 'n%'. (Ex: `{x:'10%'}` places number 10% from left edge) |
| `y`          | number  | inches | `90%`   | vertical location   | 0-n OR 'n%'. (Ex: `{y:'90%'}` places number 90% down the Slide) |
| `w`          | number  | inches | `0.3`   | containing shape width  | 0-n. Ex: `{w:1.5}` |
| `h`          | number  | inches | `0.5`   | containing shape height | 0-n. Ex: `{h:1.0}` |
| `color`      | string  |        |         | text color          | hex color code or [scheme color constant](#scheme-colors). Ex: `{color:'0088CC'}` |
| `fontFace`   | string  |        |         | font face           | any available font. Ex: `{fontFace:Arial}` |
| `fontSize`   | number  | points |         | font size           | 8-256. Ex: `{fontSize:12}` |

### Slide Number Examples
```javascript
// EX: Add a Slide Number at a given location
slide.slideNumber({ x:1.0, y:'90%' });

// EX: Styled Slide Numbers
slide.slideNumber({ x:'50%', y:'90%', fontFace:'Courier', fontSize:32, color:'CF0101' });
```
