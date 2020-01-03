---
id: api-images
title: Adding Images
---
## Syntax
```javascript
slide.addImage({OPTIONS});
```

## Usage
Either provide a URL location or base64 data to create an image.  
* `path` can be either a local or remote URL
* `data` is a base64 string representing an encoded image

## Supported Formats
* Image (png, jpg, svg, gif and animated gif, etc.)
* Note: SVG images are only supported in the newest version of PowerPoint or PowerPoint Online

## Image Options
| Option       | Type    | Unit   | Default  | Description         | Possible Values  |
| :----------- | :------ | :----- | :------- | :------------------ | :--------------- |
| `x`          | number  | inches | `1.0`    | horizontal location | 0-n |
| `y`          | number  | inches | `1.0`    | vertical location   | 0-n |
| `w`          | number  | inches | `1.0`    | width               | 0-n |
| `h`          | number  | inches | `1.0`    | height              | 0-n |
| `data`       | string  |        |          | image data (base64) | base64-encoded image string. (either `data` or `path` is required) |
| `hyperlink`  | string  |        |          | add hyperlink | object with `url` or `slide` (`tooltip` optional). Ex: `{ hyperlink:{url:'https://github.com'} }` |
| `path`       | string  |        |          | image path          | Same as used in an (img src="") tag. (either `data` or `path` is required) |
| `rounding`   | boolean |        | `false`  | image rounding      | Shapes an image into a circle |
| `sizing`     | object  |        |          | transforms image    | See [Image Sizing](#image-sizing) |

## Image Examples
```javascript
var pptx = new PptxGenJS();
var slide = pptx.addSlide();

// EX: Image by local URL
slide.addImage({ path:'images/chart_world_peace_near.png', x:1, y:1, w:8.0, h:4.0 });

// EX: Image from remote URL
slide.addImage({ path:'https://upload.wikimedia.org/wikipedia/en/a/a9/Example.jpg', x:1, y:1, w:6, h:4 })

// EX: Image by data (pre-encoded base64)
slide.addImage({ data:'image/png;base64,iVtDafDrBF[...]=', x:3.0, y:5.0, w:6.0, h:3.0 });

// EX: Image with Hyperlink
slide.addImage({
  x:1.0, y:1.0, w:8.0, h:4.0,
  hyperlink:{ url:'https://github.com/gitbrent/pptxgenjs', tooltip:'Visit Homepage' },
  path:'images/chart_world_peace_near.png',
});

pptx.writeFile('Demo-Images');
```

## Image Sizing
The `sizing` option provides cropping and scaling an image to a specified area. The property expects an object with the following structure:

| Property     | Type    | Unit   | Default           | Description                                   | Possible Values  |
| :----------- | :------ | :----- | :---------------- | :-------------------------------------------- | :--------------- |
| `type`       | string  |        |                   | sizing algorithm                              | `'crop'`, `'contain'` or `'cover'` |
| `w`          | number  | inches | `w` of the image  | area width                                    | 0-n |
| `h`          | number  | inches | `h` of the image  | area height                                   | 0-n |
| `x`          | number  | inches | `0`               | area horizontal position related to the image | 0-n (effective for `crop` only) |
| `y`          | number  | inches | `0`               | area vertical position related to the image   | 0-n (effective for `crop` only)|

Particular `type` values behave as follows:
* `contain` works as CSS property `background-size` — shrinks the image (ratio preserved) to the area given by `w` and `h` so that the image is completely visible. If the area's ratio differs from the image ratio, an empty space will surround the image.
* `cover` works as CSS property `background-size` — shrinks the image (ratio preserved) to the area given by `w` and `h` so that the area is completely filled. If the area's ratio differs from the image ratio, the image is centered to the area and cropped.
* `crop` cuts off a part specified by image-related coordinates `x`, `y` and size `w`, `h`.

NOTES:
* If you specify an area size larger than the image for the `contain` and `cover` type, then the image will be stretched, not shrunken.
* In case of the `crop` option, if the specified area reaches out of the image, then the covered empty space will be a part of the image.
* When the `sizing` property is used, its `w` and `h` values represent the effective image size. For example, in the following snippet, width and height of the image will both equal to 2 inches and its top-left corner will be located at [1 inch, 1 inch]:
```javascript
slide.addImage({
  path: '...', w:4, h:3, x:1, y:1,
  sizing: { type:'contain', w:2, h:2 }
});
```

## Performance Considerations
It takes CPU time to read and encode images! The more images you include and the larger they are, the more time will be consumed.
The time needed to read/encode images can be completely eliminated by pre-encoding any images (see below).

## Pre-Encode Large Images
Pre-encode images into a base64 string (eg: 'image/png;base64,iVBORw[...]=') for use as the `data` option value.
This will both reduce dependencies (who needs another image asset to keep track of?) and provide a performance
boost (no time will need to be consumed reading and encoding the image).
