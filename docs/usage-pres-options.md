---
id: usage-pres-options
title: Presentation Options
---

## Metadata
**************************************************************************************************

### Metadata Properties
There are several optional PowerPoint metadata properties that can be set.

### Metadata Properties Examples
```javascript
pptx.setAuthor('Brent Ely');
pptx.setCompany('S.T.A.R. Laboratories');
pptx.setRevision('15');
pptx.setSubject('Annual Report');
pptx.setTitle('PptxGenJS Sample Presentation');
```





## Slide Layouts (Sizes)
**************************************************************************************************

### Slide Layout Syntax
```javascript
pptx.setLayout('LAYOUT_NAME');
```

* Note: Layout Options apply to all the Slides in the current Presentation.

### Slide Layout Options
| Layout Name    | Default  | Layout Slide Size                 |
| :------------- | :------- | :-------------------------------- |
| `LAYOUT_16x9`  | Yes      | 10 x 5.625 inches                 |
| `LAYOUT_16x10` | No       | 10 x 6.25 inches                  |
| `LAYOUT_4x3`   | No       | 10 x 7.5 inches                   |
| `LAYOUT_WIDE`  | No       | 13.3 x 7.5 inches                 |
| `LAYOUT_USER`  | No       | user defined - see below (inches) |

Custom user defined Layout sizes are supported - just supply a `name` and the size in inches.
* Defining a new Layout using an object will also set this new size as the current Presentation Layout

### Slide Layout Examples
```javascript
// Defines and sets this new layout for the Presentation
pptx.setLayout({ name:'A3', width:16.5, height:11.7 });
```



## Text Direction
**************************************************************************************************

### Text Direction Options
Right-to-Left (RTL) text is supported.  Simply set the RTL mode at the Presentation-level.

### Text Direction Examples
```javascript
// Set right-to-left text
pptx.setRTL(true);
```
