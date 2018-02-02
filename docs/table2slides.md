---
id: table2slides
title: Table-to-Slides
---
**************************************************************************************************
Table of Contents
- [Syntax](#syntax)
- [Table-to-Slides Options](#table-to-slides-options)
- [Table-to-Slides HTML Options](#table-to-slides-html-options)
- [Table-to-Slides Notes](#table-to-slides-notes)
- [Table-to-Slides Examples](#table-to-slides-examples)
- [Creative Solutions](#creative-solutions)
**************************************************************************************************

## Table-to-Slides Syntax
```javascript
slide.addSlidesForTable(htmlElementID);
slide.addSlidesForTable(htmlElementID, {OPTIONS});
```

Any variety of HTML tables can be turned into a series of slides (auto-paging) by providing the table's ID.
* Reproduces an HTML table - background colors, borders, fonts, padding, etc.
* Slide margins are based on either the Master Slide provided or options

*NOTE: Nested tables are not supported in PowerPoint, so only the string contents of a single level deep table cell will be reproduced*

## Table-to-Slides Options
| Option            | Type    | Unit   | Description                     | Possible Values  |
| :---------------- | :------ | :----- | :------------------------------ | :--------------------------------------------- |
| `x`               | number  | inches | horizontal location             | 0-256. Table will be placed here on each Slide |
| `y`               | number  | inches | vertical location               | 0-256. Table will be placed here on each Slide |
| `w`               | number  | inches | width                           | 0-256. Default is (100% - Slide margins)       |
| `h`               | number  | inches | height                          | 0-256. Default is (100% - Slide margins)       |
| `master`          | string  |        | master slide to use             | [Slide Masters](#slide-masters) name. Ex: `{ master:'TITLE_SLIDE' }` |
| `addHeaderToEach` | boolean |        | add table headers to each slide | Ex: `addHeaderToEach:true`   |
| `addImage`        | string  |        | add an image to each slide      | Ex: `{ addImage:{ path:"images/logo.png", x:10, y:0.5, w:1.2, h:0.75 } }` |
| `addShape`        | string  |        | add a shape to each slide       | Use the established syntax   |
| `addTable`        | string  |        | add a table to each slide       | Use the established syntax   |
| `addText`         | string  |        | add text to each slide          | Use the established syntax   |

## Table-to-Slides HTML Options
Add an `data` attribute to the table's `<th>` tag to manually size columns (inches)
* minimum column width can be specified by using the `data-pptx-min-width` attribute
* fixed column width can be specified by using the `data-pptx-width` attribute

Example:
```HTML
<table id="tabAutoPaging" class="tabCool">
  <thead>
    <tr>
      <th data-pptx-min-width="0.6" style="width: 5%">Row</th>
      <th data-pptx-min-width="0.8" style="width:10%">Last Name</th>
      <th data-pptx-min-width="0.8" style="width:10%">First Name</th>
      <th data-pptx-width="8.5"     style="width:75%">Description</th>
    </tr>
  </thead>
  <tbody></tbody>
</table>
```

## Table-to-Slides Notes
* Default `x`, `y` and `margin` value is 0.5 inches, the table will take up all remaining space by default (h:100%, w:100%)
* Your Master Slides should already have defined margins, so a Master Slide name is the only option you'll need most of the time
* Hidden tables wont auto-size their columns correctly (as the properties are not accurate)

## Table-to-Slides Examples
```javascript
// Pass table element ID to addSlidesForTable function to produce 1-N slides
pptx.addSlidesForTable( 'myHtmlTableID' );

// Optionally, include a Master Slide name for pre-defined margins, background, logo, etc.
pptx.addSlidesForTable( 'myHtmlTableID', { master:'MASTER_SLIDE' } );

// Optionally, add images/shapes/text/tables to each Slide
pptx.addSlidesForTable( 'myHtmlTableID', { addText:{ text:"Dynamic Title", options:{x:1, y:0.5, color:'0088CC'} } } );
pptx.addSlidesForTable( 'myHtmlTableID', { addImage:{ path:"images/logo.png", x:10, y:0.5, w:1.2, h:0.75 } } );
```

## Creative Solutions
Design a Master Slide that already contains: slide layout, margins, logos, etc., then you can produce
professional looking Presentations with a single line of code which can be embedded into a link or a button:

Add a button to a webpage that will create a Presentation using whatever table data is present:
```html
<input type="button" value="Export to PPTX" onclick="{ var pptx=new PptxGenJS(); pptx.addSlidesForTable('tableId'); pptx.save(); }">
```

**SharePoint Integration**

Placing a button like this into a WebPart is a great way to add "Export to PowerPoint" functionality
to SharePoint. (You'd also need to add the 4 `<script>` includes in the same or another WebPart)
