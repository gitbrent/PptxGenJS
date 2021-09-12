---
id: speaker-notes
title: Speaker Notes
---

Speaker Notes can be included on any Slide.

## Syntax
```typescript
slide.addNotes('TEXT');
```

## Example: JavaScript
```typescript
let pres = new PptxGenJS();
let slide = pptx.addSlide();

slide.addText('Hello World!', { x:1.5, y:1.5, fontSize:18, color:'363636' });

slide.addNotes('This is my favorite slide!');

pptx.writeFile('Sample Speaker Notes');
```

## Example: TypeScript
```typescript
import pptxgen from "pptxgenjs";

let pres = new pptxgen();
let slide = pptx.addSlide();

slide.addText('Hello World!', { x:1.5, y:1.5, fontSize:18, color:'363636' });

slide.addNotes('This is my favorite slide!');

pptx.writeFile('Sample Speaker Notes');
```
