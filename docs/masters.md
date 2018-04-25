---
id: masters
title: Master Slides
---

## Slide Masters
Generating sample slides like those shown in the Examples section are great for demonstrating library features,
but the reality is most of us will be required to produce presentations that have a certain design or
corporate branding.

PptxGenJS allows you to define Slide Master Layouts via objects that can then be used to provide branding
functionality.  This enables you to easily create a Master Slide using code.

Slide Masters are created by calling the `defineSlideMaster()` method along with an options object
(same style used in Slides).  Once defined, you can pass the Master title to `addNewSlide()` and that Slide will
use the Layout previously defined.  See the demo under /examples for several working examples.

The defined Masters become first-class Layouts in the exported PowerPoint presentation and can be changed
via View > Slide Master and will affect the Slides created using that layout.

## Slide Master Options
| Option        | Type    | Unit   | Default  | Description  | Possible Values       |
| :------------ | :------ | :----- | :------- | :----------- | :-------------------- |
| `bkgd`        | string  |        | `ffffff` | color        | hex color code or [scheme color constant](#scheme-colors). Ex: `{ bkgd:'0088CC' }` |
| `bkgd`        | object  |        |          | image | object with path OR data. Ex: `{path:'img/bkgd.png'}` OR `{data:'image/png;base64,iVBORwTwB[...]='}` |
| `slideNumber` | object  |        |          | Show slide numbers | ex: `{ x:1.0, y:'50%' }` `x` and `y` can be either inches or percent |
| `margin`      | number  | inches | `1.0`    | Slide margins      | 0.0 through Slide.width |
| `margin`      | array   |        |          | Slide margins      | array of numbers in TRBL order. Ex: `[0.5, 0.75, 0.5, 0.75]` |
| `objects`     | array   |        |          | Objects for Slide  | object with type and options. Type:`chart`,`image`,`line`,`rect` or `text`. [Example](https://github.com/gitbrent/PptxGenJS#slide-master-examples) |
| `title`       | string  |        |          | Layout title/name  | some title |

**TIP:**
Pre-encode your images (base64) and add the string as the optional data key/val (see `bkgd` above)

## Slide Master Examples
```javascript
var pptx = new PptxGenJS();
pptx.setLayout('LAYOUT_WIDE');

pptx.defineSlideMaster({
  title: 'MASTER_SLIDE',
  bkgd:  'FFFFFF',
  objects: [
    { 'line':  { x: 3.5, y:1.00, w:6.00, line:'0088CC', lineSize:5 } },
    { 'rect':  { x: 0.0, y:5.30, w:'100%', h:0.75, fill:'F1F1F1' } },
    { 'text':  { text:'Status Report', options:{ x:3.0, y:5.30, w:5.5, h:0.75 } } },
    { 'image': { x:11.3, y:6.40, w:1.67, h:0.75, path:'images/logo.png' } }
  ],
  slideNumber: { x:0.3, y:'90%' }
});

var slide = pptx.addNewSlide('MASTER_SLIDE');
slide.addText('How To Create PowerPoint Presentations with JavaScript', { x:0.5, y:0.7, fontSize:18 });

pptx.save();
```

### Slide Master Demo
There are several Master Slides defined in the Demo: `examples/pptxgenjs-demo.html`
![PptxGenJS Master Slide Demo](/PptxGenJS/docs/assets/ex-master-slide-demo.png)

### Slide Master Output
Using the 'MASTER_SLIDE' defined above to produce a Slide:
![Master Slide Demo Presentation](/PptxGenJS/docs/assets/ex-master-slide-output.png)
