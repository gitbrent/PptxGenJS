---
id: usage-add-slide
title: Adding a Slide
---

## Syntax
```javascript
var slide = pptx.addSlide();
```

See [Slide Options](/PptxGenJS/docs/usage-slide-options.html) for features such as Slide Numbers.

## Slide Templates/Master Slides
**************************************************************************************************

### Master Slide Syntax
```javascript
var slide = pptx.addSlide('MASTER_NAME');
```

(See [Master Slides](/PptxGenJS/docs/masters.html) for more about creating masters/templates)

### Master Slide Examples
```javascript
// Create a new Slide that will inherit properties from a pre-defined master page (margins, logos, text, background, etc.)
var slide = pptx.addSlide('TITLE_SLIDE');
```



## Default Slide Colors
**************************************************************************************************

### Default Slide Color Options
| Option       | Type    | Default   | Description         | Possible Values  |
| :----------- | :------ | :-------- | :------------------ | :--------------- |
| `bkgd`       | string  | `FFFFFF`  | background color    | hex color code or [scheme color constant](#scheme-colors). |
| `color`      | string  | `000000`  | default text color  | hex color code or [scheme color constant](#scheme-colors). |

### Default Slide Color Examples
```javascript
// Set slide background color
slide.bkgd = 'F1F1F1';

// Set slide default font color
slide.color = '696969';
```



## Slides Return Themselves
**************************************************************************************************
The Slide object returns a reference to itself, so calls can be chained.

Example:
```javascript
slide.addImage({ path:'img1.png', x:1, y:2 }).addImage({ path:'img2.jpg', x:5, y:3 });
```
