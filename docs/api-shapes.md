---
id: api-shapes
title: Adding Shapes
---
## Syntax

### Plain/no text
```javascript
slide.addShape({SHAPE}, {OPTIONS});
```

### With text
```javascript
slide.addText("some string", {SHAPE, OPTIONS});
```
Check the `pptxgen.shapes.js` file for a complete list of the hundreds of PowerPoint shapes available.

## Shape Options
| Option       | Type    | Unit   | Default | Description         | Possible Values                                                             |
| :----------- | :------ | :----- | :------ | :------------------ | :-------------------------------------------------------------------------- |
| `x`          | number  | inches | `1.0`   | horizontal location | 0-n OR 'n%'. (Ex: `{x:'50%'}` will place object in the middle of the Slide) |
| `y`          | number  | inches | `1.0`   | vertical location   | 0-n OR 'n%'.                                                                |
| `w`          | number  | inches |         | width               | 0-n OR 'n%'. (Ex: `{w:'50%'}` will make object 50% width of the Slide)      |
| `h`          | number  | inches |         | height              | 0-n OR 'n%'.                                                                |
| `align`      | string  |        | `left`  | alignment           | `left` or `center` or `right`                                               |
| `fill`       | string  |        |         | fill/bkgd color     | hex color code or [scheme color constant](#scheme-colors). Ex: `{color:'0088CC'}` |
| `fill`       | object  |        |         | fill/bkgd color  | object with `type`, `color`, `alpha` (opt). Ex: `fill:{type:'solid', color:'0088CC', alpha:25}` |
| `flipH`      | boolean |        |         | flip Horizontal     | `true` or `false` |
| `flipV`      | boolean |        |         | flip Vertical       | `true` or `false` |
| `line`       | string  |        |         | border line color   | hex color code or [scheme color constant](#scheme-colors). Ex: `{line:'0088CC'}` |
| `lineDash`   | string  |        | `solid` | border line dash style | `dash`, `dashDot`, `lgDash`, `lgDashDot`, `lgDashDotDot`, `solid`, `sysDash` or `sysDot` |
| `line_head`   | string  |        |         | border line ending  | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none` |
| `lineSize`   | number  | points |         | border line size    | 1-256. Ex: {lineSize:4} |
| `line_tail`  | string  |        |         | border line heading | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none` |
| `rectRadius` | number  | inches |         | rounding radius     | rounding radius for `ROUNDED_RECTANGLE` text shapes |
| `rotate`     | integer | degrees|         | rotation degrees    | 0-360. Ex: `{rotate:180}` |

## Shape Examples
```javascript
var pptx = new PptxGenJS();
pptx.setLayout('LAYOUT_WIDE');

var slide = pptx.addNewSlide();

// LINE
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:4.40, w:5, h:0, line:'FF0000', lineSize:1 });
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:4.80, w:5, h:0, line:'FF0000', lineSize:2, line_head:'triangle' });
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:5.20, w:5, h:0, line:'FF0000', lineSize:3, line_tail:'triangle' });
slide.addShape(pptx.shapes.LINE,      { x:4.15, y:5.60, w:5, h:0, line:'FF0000', lineSize:4, line_head:'triangle', line_tail:'triangle' });

// DIAGONAL LINE
slide.addShape(pptx.shapes.LINE,      { x:0, y:0, w:5.0, h:0, line:'FF0000', rotate:45 });

// RECTANGLE
slide.addShape(pptx.shapes.RECTANGLE, { x:0.50, y:0.75, w:5, h:3, fill:'FF0000' });

// OVAL
slide.addShape(pptx.shapes.OVAL,      { x:4.15, y:0.75, w:5, h:2, fill:{ type:'solid', color:'0088CC', alpha:25 } });

// Adding text to Shapes:
slide.addText('RIGHT-TRIANGLE', { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:0.40, y:4.3, w:6, h:3, fill:'0088CC', line:'000000', lineSize:3 });
slide.addText('RIGHT-TRIANGLE', { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:7.00, y:4.3, w:6, h:3, fill:'0088CC', line:'000000', flipH:true });

pptx.save('Demo-Shapes');
```
