---
id: usage-saving
title: Saving Presentations
---

Presentations require nothing more than passing a filename to `save()`. Node.js users have more options available
examples of which can be found below.

**************************************************************************************************
## Client Browser
* Simply provide a filename

```javascript
pptx.save('Demo-Media');
```

**************************************************************************************************
## Node.js
* Node can accept a callback function that will return the filename once the save is complete
* Node can also be used to stream a powerpoint file - simply pass a filename that begins with "http"
* Output type can be specified by passing an optional [JSZip output type](https://stuk.github.io/jszip/documentation/api_jszip/generate_async.html)

```javascript
// Example A: File will be saved to the local working directory (`__dirname`)
pptx.save('Node_Demo');

// Example B: Inline callback function
pptx.save('Node_Demo', function(filename){ console.log('Created: '+filename); });

// Example C: Predefined callback function
pptx.save('Node_Demo', saveCallback);

// Example D: Use a filename of "http" or "https" to receive the powerpoint binary data in your callback
// (Used for streaming the presentation file via http.  See the `nodejs-demo.js` file for a working example.)
pptx.save('http', streamCallback);

// Example E: Save using various JSZip output types: ['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array']
pptx.save('jszip', saveCallback, 'base64');
```

**************************************************************************************************
## Saving Multiple Presentations

### Client-Side:
* In order to generate a new, unique Presentation just create a new instance of the library then add objects and save as normal.

```javascript
var pptx = null;

// Presentation 1:
pptx = new PptxGenJS();
pptx.addNewSlide().addText('Presentation 1', {x:1, y:1});
pptx.save('PptxGenJS-Presentation-1');

// Presentation 2:
pptx = new PptxGenJS();
pptx.addNewSlide().addText('Presentation 2', {x:1, y:1});
pptx.save('PptxGenJS-Presentation-2');
```

### Node.js:
* See `examples/nodejs-demo.js` for a working demo with multiple presentations, callbacks, streaming, etc.

```javascript
var PptxGenJS = require("pptxgenjs");
var pptx = null;

// Presentation 1:
pptx = new PptxGenJS();
pptx.addNewSlide().addText('Presentation 1', {x:1, y:1});
pptx.save('PptxGenJS-NodePres-1');

// Presentation 2:
pptx = new PptxGenJS();
pptx.addNewSlide().addText('Presentation 2', {x:1, y:1});
pptx.save('PptxGenJS-NodePres-2');
```
