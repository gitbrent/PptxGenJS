# PptxGenJS
A complete PowerPoint/OpenOffice/Slides Presentation creator framework for client browsers.

Supported write formats:
* PowerPoint 2007+ XML Formats (PPTX)

Demo: [http://gitbrent.github.io/PptxGenJS](http://gitbrent.github.io/PptxGenJS)

Source: [http://git.io/pptxgenjs](http://git.io/pptxgenjs)

# Installation

```javascript
<script lang="javascript" src="dist/filesaver.js"></script>
<script lang="javascript" src="dist/jquery.min.js"></script>
<script lang="javascript" src="dist/jszip.min.js"></script>
<script lang="javascript" src="dist/pptxgen.min.js"></script>
```

# Optional Modules
```javascript
<script lang="javascript" src="dist/pptxgen.masters.js"></script>
<script lang="javascript" src="dist/pptxgen.shapes.js"></script>
```

# Writing Presentation
Creating a Presentation is super easy.

```javascript
var pptx = new PptxGenJS();
pptx.setLayout('LAYOUT_16x9');
slide.addText('Hello World', { x:0.5, y:0.7, font_size:18, font_face:'Arial', color:'0000FF' });
pptx.doExportPptx('HelloWorld');
```

# Lastly
NOTE: It takes time to encode images, so the more images you include and the larger they are, the more time will be consumed.
You will want to show a jQuery Dialog with a nice hour glass before you start creating the file.
