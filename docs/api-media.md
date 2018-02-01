---
id: api-media
title: Adding Media
---
**************************************************************************************************
Table of Contents
- [Syntax](#syntax)
- [Media Options](#media-options)
- [Media Examples](#media-examples)
**************************************************************************************************

## Syntax
```javascript
slide.addMedia({OPTIONS});
```

Both Video (mpg, mov, mp4, m4v, etc.) and Audio (mp3, wav, etc.) are supported.
* PowerPoint [supported audio/video formats](https://support.office.com/en-us/article/Video-and-audio-file-formats-supported-in-PowerPoint-d8b12450-26db-4c7b-a5c1-593d3418fb59#OperatingSystem=Windows)
* Node.js: use either `data` or `path` options (Node can encode any media into base64)
* Browsers: pre-encode the media into base64 and add it using the `data` option (current browsers cannot encode media)

Online video (YouTube embeds, etc.) is supported in both client browser and in Node.js
* Currently, only PowerPoint Online can display YouTube videos

## Media Options
| Option       | Type    | Unit   | Default   | Description         | Possible Values  |
| :----------- | :------ | :----- | :-------- | :------------------ | :--------------- |
| `x`          | number  | inches | `1.0`     | horizontal location | 0-n |
| `y`          | number  | inches | `1.0`     | vertical location   | 0-n |
| `w`          | number  | inches | `1.0`     | width               | 0-n |
| `h`          | number  | inches | `1.0`     | height              | 0-n |
| `data`       | string  |        |           | media data (base64) | base64-encoded string |
| `path`       | string  |        |           | media path          | relative path to media file |
| `link`       | string  |        |           | online url/link     | link to online video. Ex: `link:'https://www.youtube.com/embed/blahBlah'` |
| `type`       | string  |        |           | media type          | media type: `audio` or `video` (reqs: `data` or `path`) or `online` (reqs:`link`) |

## Media Examples
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();

// Media by path (Node.js only)
slide.addMedia({ type:'audio', path:'../media/sample.mp3', x:1.0, y:1.0, w:3.0, h:0.5 });
// Media by data (client browser or Node.js)
slide.addMedia({ type:'audio', data:'audio/mp3;base64,iVtDafDrBF[...]=', x:3.0, y:1.0, w:6.0, h:3.0 });
// Online by link (client browser or Node.js)
slide.addMedia({ type:'online', link:'https://www.youtube.com/embed/Dph6ynRVyUc', x:1.0, y:4.0, w:8.0, h:4.5 });

pptx.save('Demo-Media');
```
