---
id: quick-start
title: Quick Start Guide
---

## Library Support
* Include via `<script>` for web applications
* npm/yarn installation for Angular, React, Electron, or NodeJS


## Creating a Presentation
PptxGenJS PowerPoint presentations are created via JavaScript by following 4 basic steps:

1. Create a Presentation
2. Add a Slide
3. Add an object (Chart, Shape, Table, etc.) to the Slide
4. Save the Presentation

### Simple Example
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addSlide();
slide.addText('Hello World!', { x:1.5, y:1.5, fontSize:18, color:'363636' });
pptx.writeFile('Sample Presentation');
```
That's really all there is to it!


## TypeScript Support
If you're using Angular or React, the included TypeScript definitions file brings the documentation to you.

[[TODO:show screen cap]]
