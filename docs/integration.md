---
id: integration
title: Library Integration
---

## Integration with Angular/React

* Working React demo under [demos/react-demo](https://github.com/gitbrent/PptxGenJS/tree/master/demos/react-demo)

### React Example  
```
import pptxgen from "pptxgenjs";

let pptx = new pptxgen();
let slide = pptx.addSlide();

slide.addText("React Demo!", \{ x:1, y:1, w:'80%', h:1, fontSize:36, fill:'eeeeee', align:'center' \});
pptx.writeFile("react-demo.pptx");
```


## Integration with Webpack/TypeScript

Some users have modified their webpack config to avoid a module resolution error using:
* `node: { fs: "empty" }`  

## More Information
- [See Issue #72 for more information](https://github.com/gitbrent/PptxGenJS/issues/72)
- [See Issue #220 for more information](https://github.com/gitbrent/PptxGenJS/issues/220)
- [See Issue #308 for more information](https://github.com/gitbrent/PptxGenJS/issues/308)


## Available Library Distributions
- CommonJS `dist/pptxgen.cjs.js`
- ES6 Module `dist/pptxgen.es.js`
