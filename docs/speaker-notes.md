---
id: speaker-notes
title: Speaker Notes
---
Add Speaker Notes to any slide.

## Syntax
```javascript
slide.addNotes('TEXT');
```

## Example
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addSlide();

slide.addText('Hello World!', { x:1.5, y:1.5, fontSize:18, color:'363636' });

slide.addNotes('This is my favorite slide!');

pptx.writeFile('Sample Speaker Notes');
```
