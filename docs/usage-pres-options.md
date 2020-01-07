---
id: usage-pres-options
title: Presentation Options
---

## Metadata
**************************************************************************************************

### Metadata Properties
There are several optional PowerPoint metadata properties that can be set.

### Metadata Properties Examples
PptxGenJS uses ES6-style getters/setters.

```javascript
pptx.author = 'Brent Ely';
pptx.company = 'S.T.A.R. Laboratories';
pptx.revision = '15';
pptx.subject = 'Annual Report';
pptx.title = 'PptxGenJS Sample Presentation';
```



## Slide Layouts (Sizes)
**************************************************************************************************
Layout Option applies to all the Slides in the current Presentation.

### Slide Layout Syntax
```javascript
pptx.layout = 'LAYOUT_NAME';
```

### Standard Slide Layouts
| Layout Name    | Default  | Layout Slide Size   |
| :------------- | :------- | :------------------ |
| `LAYOUT_16x9`  | Yes      | 10 x 5.625 inches   |
| `LAYOUT_16x10` | No       | 10 x 6.25 inches    |
| `LAYOUT_4x3`   | No       | 10 x 7.5 inches     |
| `LAYOUT_WIDE`  | No       | 13.3 x 7.5 inches   |

### Custom Slide Layouts
Custom, user-defined layouts are supported
* Use the `defineLayout()` method to create a custom layout of any size
* Create as many layouts as needed, ex: create an 'A3' and 'A4' and set layouts as desired

### Custom Slide Layout Example
```javascript
// Define new layout for the Presentation
pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });

// Set presentation to use new layout
pptx.layout = 'A3'
```



## Text Direction
**************************************************************************************************

### Text Direction Options
Right-to-Left (RTL) text is supported. Simply set the RTL mode at the Presentation-level.

### Text Direction Examples
```javascript
// Set right-to-left text mode
pptx.rtlMode = true;
```

Notes:
* You may also need to set an RTL lang value such as `lang='he'` as the default lang is 'EN-US'
* See [Issue#600](https://github.com/gitbrent/PptxGenJS/issues/600) for more
