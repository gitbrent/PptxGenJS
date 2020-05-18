---
id: api-shapes
title: Adding Shapes
---
## Syntax

### Shape without text
```javascript
slide.addShape( SHAPE, OPTIONS );
```

### Shape with text
```javascript
slide.addText( "some string", {SHAPE, OPTIONS} );
```


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
| `lineHead`   | string  |        |         | border line ending  | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none` |
| `lineSize`   | number  | points |         | border line size    | 1-256. Ex: {lineSize:4} |
| `lineTail`   | string  |        |         | border line heading | `arrow`, `diamond`, `oval`, `stealth`, `triangle` or `none` |
| `rectRadius` | number  | inches |         | rounding radius     | rounding radius for `ROUNDED_RECTANGLE` text shapes |
| `rotate`     | integer | degrees|         | rotation degrees    | 0-360. Ex: `{rotate:180}` |

## Shape Examples
```javascript
// Plain shapes:
let slide1 = pres.addSlide();
slide1.addShape(pres.ShapeType.rect,      { x:0.5, y:0.8, w:1.5, h:3.0, fill:'FF0000' });
slide1.addShape(pres.ShapeType.ellipse,   { x:5.4, y:0.8, w:3.0, h:1.5, fill:{ type:'solid', color:'0088CC' } });
slide1.addShape(pres.ShapeType.line,      { x:4.2, y:4.4, w:5.0, h:0.0, line:'FF0000', lineSize:1 });
slide1.addShape(pres.ShapeType.line,      { x:4.2, y:4.8, w:5.0, h:0.0, line:'FF0000', lineSize:2, lineHead:'triangle' });

// Shapes with Text:
let slide2 = pres.addSlide();
slide2.addText('RECTANGLE',   { shape:pres.ShapeType.rect, x:0.5, y:0.8, w:1.5, h:3.0, fill:'FF0000', align:'center', fontSize:14 });
slide2.addText('ELLIPSE',     { shape:pres.ShapeType.ellipse,      x:5.4, y:0.8, w:3.0, h:1.5, fill:'F38E00', align:'center', fontSize:14 });
slide2.addText('LINE size=1', { shape:pres.ShapeType.line, align:'center', x:4.2, y:4.4, w:5, h:0, line:'FF0000', lineSize:1, lineDash:'lgDash' });
slide2.addText('LINE size=2', { shape:pres.ShapeType.line, align:'left',   x:4.2, y:4.8, w:5, h:0, line:'FF0000', lineSize:2, lineTail:'triangle' });

pres.writeFile('Demo-Shapes');
```
