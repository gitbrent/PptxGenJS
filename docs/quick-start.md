---
id: quick-start
title: Quick Start Guide
---

## Library Support
* Integrates with Angular, React and Electron

## Creating a Presentation
PptxGenJS PowerPoint presentations are created via JavaScript by following 4 basic steps:

1. Create a new Presentation
2. Add a Slide
3. Add one or more objects (Tables, Shapes, Images, Text and Media) to the Slide
4. Save the Presentation

### Code Example
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();
slide.addText('Hello World!', { x:1.5, y:1.5, fontSize:18, color:'363636' });
pptx.save('Sample Presentation');
```

### Presentation Done
That's really all there is to it!
