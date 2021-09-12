---
id: integration
title: Integration
---

## Available Distributions

-   Browser `dist/pptxgen.min.js`
-   CommonJS `dist/pptxgen.cjs.js`
-   ES6 Module `dist/pptxgen.es.js`

## Integration with Angular/React

-   There is a working demo available: [demos/react-demo](https://github.com/gitbrent/PptxGenJS/tree/master/demos/react-demo)

### React Example

```typescript
import pptxgen from "pptxgenjs";
let pptx = new pptxgen();

let slide = pptx.addSlide();
slide.addText("React Demo!", { x: 1, y: 1, w: 10, fontSize: 36, fill: { color: "F1F1F1" }, align: "center" });

pptx.writeFile({ fileName: "react-demo.pptx" });
```

## Webpack Troubleshooting

Some users have modified their webpack config to avoid a module resolution error using:

-   `node: { fs: "empty" }`

### Related Issues

-   [See Issue #72 for more information](https://github.com/gitbrent/PptxGenJS/issues/72)
-   [See Issue #220 for more information](https://github.com/gitbrent/PptxGenJS/issues/220)
-   [See Issue #308 for more information](https://github.com/gitbrent/PptxGenJS/issues/308)
