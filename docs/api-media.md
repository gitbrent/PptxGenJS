---
id: api-media
title: Adding Media
---

## Syntax
```javascript
slide.addMedia({OPTIONS});
```

## Usage
Either provide a URL location or base64 data to create media.  
* `path` can be either a local or remote URL
* `data` is a base64 string representing an encoded media (hit-or-miss situation based upon recent feedback)

## Supported Formats
* Video (mpg, mov, mp4, m4v, etc.)
* Audio (mp3, wav, etc.)
* (Reference: [Video and Audio file formats supported in PowerPoint](https://support.office.com/en-us/article/Video-and-audio-file-formats-supported-in-PowerPoint-d8b12450-26db-4c7b-a5c1-593d3418fb59#OperatingSystem=Windows))

## Media Notes
* Not all platforms support all formats! MacOS can show MPG files whereas Windows probably will not, and some AVI
files may work and some may not.  Video codecs are weird and painful like that.
* YouTube videos work great in Microsoft Office online... other video sites, not so much (YMMV).

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

### Media Examples
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addNewSlide();

// EX: Media by path
slide.addMedia({ type:'video', path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/2.1.0/examples/media/sample.mov', x:1.0, y:1.0, w:3.0, h:2.0 });
slide.addMedia({ type:'audio', path:'../media/sample.mp3', x:1.0, y:1.0, w:3.0, h:0.5 });

// EX: Media by data (does not always work well - use URL instead)
slide.addMedia({ type:'audio', data:'audio/mp3;base64,iVtDafDrBF[...]=', x:3.0, y:1.0, w:6.0, h:3.0 });

// EX: YouTube video
slide.addMedia({ type:'online', link:'https://www.youtube.com/embed/Dph6ynRVyUc', x:1.0, y:4.0, w:8.0, h:4.5 });

pptx.save('Demo-Media');
```
