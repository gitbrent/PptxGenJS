---
id: usage-pres-options
title: Presentation Options
---
**************************************************************************************************
Table of Contents
- [Presentation Properties](#presentation-properties)
- [Presentation Layouts](#presentation-layouts)
- [Presentation Layout Options](#presentation-layout-options)
- [Presentation Text Direction](#presentation-text-direction)
**************************************************************************************************

A "Presentation" is a single `.pptx` file.  See [multiple presentations](/PptxGenJS/docs/usage-saving.html#saving-multiple-presentations) for information
on creating more than a one PPT file at a time.

## New Presenation: Client Browser
```javascript
var pptx = new PptxGenJS();
```

## New Presenation: Node.js
```javascript
var PptxGenJS = require("pptxgenjs");
var pptx = new PptxGenJS();
```

## Presentation Properties
There are several optional PowerPoint metadata properties that can be set:

```javascript
pptx.setAuthor('Brent Ely');
pptx.setCompany('S.T.A.R. Laboratories');
pptx.setRevision('15');
pptx.setSubject('Annual Report');
pptx.setTitle('PptxGenJS Sample Presentation');
```

## Presentation Layouts
Setting the Layout (applies to all Slides in the Presentation):
```javascript
pptx.setLayout('LAYOUT_WIDE');
```

## Presentation Layout Options
| Layout Name    | Default  | Layout Slide Size                 |
| :------------- | :------- | :-------------------------------- |
| `LAYOUT_16x9`  | Yes      | 10 x 5.625 inches                 |
| `LAYOUT_16x10` | No       | 10 x 6.25 inches                  |
| `LAYOUT_4x3`   | No       | 10 x 7.5 inches                   |
| `LAYOUT_WIDE`  | No       | 13.3 x 7.5 inches                 |
| `LAYOUT_USER`  | No       | user defined - see below (inches) |

Custom user defined Layout sizes are supported - just supply a `name` and the size in inches.
* Defining a new Layout using an object will also set this new size as the current Presentation Layout

```javascript
// Defines and sets this new layout for the Presentation
pptx.setLayout({ name:'A3', width:16.5, height:11.7 });
```

## Presentation Text Direction
Right-to-Left (RTL) text is supported.  Simply set the RTL mode at the Presentation-level.
```javascript
pptx.setRTL(true);
```
