---
id: api-media
title: Adding Media
---
**************************************************************************************************
Table of Contents
- [Syntax](#syntax)
- [Supported Formats](#supported-formats)
- [Video Notes](#video-notes)
- [Data Options](#data-options)
- [Media Options](#media-options)
- [Media Examples](#media-examples)
**************************************************************************************************

## Syntax
```javascript
slide.addMedia({OPTIONS});
```

**IMPORTANT NOTE:**  
Adding media is predominately a Node.js feature. Why? Because no web browser can encode media files
into base64, which is the format needed to create the PPTX export file.

Support for [Adding Images](/PptxGenJS/docs/api-images.html) can be accomplished in browsers because a shadow canvas element
is created, filled using an image path, and then converted to base64 using a built-in canvas method.  No
such methods exist for media, hence, the inability to support this functionality outside of Node.

You can try to pre-encode media into base64 and pass it using the `data` option, but it is a
hit-or-miss situation based upon recent feedback.

## Supported Formats
* Video (mpg, mov, mp4, m4v, etc.)
* Audio (mp3, wav, etc.)
* (Here are the Microsoft Office [supported Audio and Video formats](https://support.office.com/en-us/article/Video-and-audio-file-formats-supported-in-PowerPoint-d8b12450-26db-4c7b-a5c1-593d3418fb59#OperatingSystem=Windows))

## Video Notes
* YouTube works great in Microsoft Office online.  Other video sites... not so much (YMMV).
* Online video linked to in the presentation (YouTube, etc.) is supported in both client browser and in Node.js
* Not all platforms support all formats! MacOS can show MPG files whereas Windows probably will not, and some AVI
files may work and some may not.  Video codecs are weird and painful like that.

## Data Options
* Node.js: use either `data` or `path` options (Node can encode any media into base64)
* Browsers: pre-encode the media and add it using the `data` option (this may not always work for various reasons)

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
