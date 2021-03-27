---
id: usage-saving
title: Saving Presentations
---

Several methods are available when generating a presentation.

- All methods return a Promise
- Working examples are available under [/demos](https://github.com/gitbrent/PptxGenJS/tree/master/demos)

## Export Methods

### Regular File

Export the presentation as a regular flat file. Browser-based apps will prompt the user to download the file and push a
PowerPoint Presentation MIME-type pptx file. Node.js apps will use `fs` to save the file to the local file system.

#### Example
```javascript
// For simple cases, you can omit `then`
pptx.writeFile('Browser-PowerPoint-Demo.pptx');

// Using Promise to determine when the file has actually completed generating
pptx.writeFile('Browser-PowerPoint-Demo.pptx');
    .then(fileName => {
        console.log(`created file: ${fileName}`);
    });
```


### Blob and Other Types
Export the presentation into other data formats for cloud storage, etc.

* Available formats: `arraybuffer`, `base64`, `binarystring`, `blob`, `nodebuffer`, `uint8array`

#### Example
```javascript
pptx.write('base64')
    .then(data => {
        console.log("write as base64: Here are 0-100 chars of `data`:\n");
        console.log(data.substring(0, 100));
    })
    .catch(err => {
        console.error(err);
    });
```


### Stream
Export the presentation into a binary stream format for cloud storage, etc.

#### Example
```javascript
// Ex using: `const app = express();``
pptx.stream()
    .then(data => {
        app.get("/", (req, res) => {
            res.writeHead(200, { "Content-disposition": "attachment;filename=" + fileName, "Content-Length": data.length });
            res.end(new Buffer(data, "binary"));
        });

        app.listen(3000, () => {
            console.log("PptxGenJS Node Stream Demo app listening on port 3000!");
            console.log("Visit: http://localhost:3000/");
            console.log("(press Ctrl-C to quit demo)");
        });
    })
    .catch(err => {
        console.log("ERROR: " + err);
    });
```



## Saving Multiple Presentations

### Client Browser
- In order to generate a new, unique Presentation just create a new instance of the library then add objects and save as normal.

```javascript
let pptx = null;

// Presentation 1:
pptx = new PptxGenJS();
pptx.addSlide().addText('Presentation 1', {x:1, y:1});
pptx.writeFile('PptxGenJS-Presentation-1');

// Presentation 2:
pptx = new PptxGenJS();
pptx.addSlide().addText('Presentation 2', {x:1, y:1});
pptx.writeFile('PptxGenJS-Presentation-2');
```

### Node.js
- See `demos/node/demo.js` for a working demo with multiple presentations, promises, etc.
- See `demos/node/demo_stream.js` for a working demo using streaming

```javascript
let PptxGenJS = require("pptxgenjs");

// Presentation 1:
let pptx1 = new PptxGenJS();
pptx1.addSlide().addText('Presentation 1', {x:1, y:1});
pptx1.writeFile('PptxGenJS-NodePres-1');

// Presentation 2:
let pptx2 = new PptxGenJS();
pptx2.addSlide().addText('Presentation 2', {x:1, y:1});
pptx2.writeFile('PptxGenJS-NodePres-2');
```
