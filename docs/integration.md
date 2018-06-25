---
id: integration
title: Library Integration
---

## Integration with Angular

Set the browser mode option so the library will use `window` blob file saving instead of detecting
your app as a Node.js app (Node apps utilize `fs.writeFile` or file streaming).  
* `pptx.setBrowser(true);`

### Angular Example  
```
const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();

pptx.setBrowser(true);

export function generatePPT() {
    const slide = pptx.addNewSlide();
    const opts = { x: 1.0, y: 1.0, font_size: 42, color: '00FF00' };
    slide.addText('Hello World!', opts);
    pptx.save();
}
```

### More Information
* [See Issue #220 for more information](https://github.com/gitbrent/PptxGenJS/issues/220)
* [See Issue #308 for more information](https://github.com/gitbrent/PptxGenJS/issues/308)


## Integration with Webpack/Typescript

Add this to your webpack config to avoid a module resolution error:
* `node: { fs: "empty" }`  

Set browser mode so files will save as blobs via browser:
* `pptx.setBrowser(true);`

[See Issue #72 for more information](https://github.com/gitbrent/PptxGenJS/issues/72)
