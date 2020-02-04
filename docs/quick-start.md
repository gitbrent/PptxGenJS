---
id: quick-start
title: Quick Start Guide
---

## Library Support

- Include via `<script>` for client-browser applications
- Install via npm/yarn Angular, React, Electron, or NodeJS

## Quick Start Guide

PptxGenJS PowerPoint presentations are created via JavaScript by following 4 basic steps:

### React/Angular, ES6, TypeScript

```typescript
import pptxgen from "pptxgenjs";

// 1. Create a new Presentation
let pres = new pptxgen();

// 2. Add a Slide
let slide = pres.addSlide();

// 3. Add one or more objects (Tables, Shapes, Images, Text and Media) to the Slide
let textboxText = "Hello World from PptxGenJS!";
let textboxOpts = { x: 1, y: 1, color: "363636", fill: "f1f1f1", align: pptx.AlignH.center };
slide.addText(textboxText, textboxOpts);

// 4. Save the Presentation
pres.writeFile("Sample Presentation.pptx");
```

### Script/Web Browser

```javascript
// 1. Create a new Presentation
let pres = new PptxGenJS();

// 2. Add a Slide
let slide = pres.addSlide();

// 3. Add one or more objects (Tables, Shapes, Images, Text and Media) to the Slide
let textboxText = "Hello World from PptxGenJS!";
let textboxOpts = { x: 1, y: 1, color: "363636", fill: "f1f1f1", align: "center" };
slide.addText(textboxText, textboxOpts);

// 4. Save the Presentation
pres.writeFile("Sample Presentation.pptx");
```

That's really all there is to it!

## TypeScript Support

If you're using Angular or React, the included TypeScript definitions file brings the documentation to you.

![TypeScript Support](/PptxGenJS/docs/assets/ex-typescript.png)
