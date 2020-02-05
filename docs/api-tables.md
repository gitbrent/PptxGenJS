---
id: api-tables
title: Adding Tables
---

## Syntax

```javascript
slide.addTable( [rows] );
slide.addTable( [rows], {any Layout/Formatting OPTIONS} );
```

## Table Layout Options

| Option | Type    | Unit   | Default | Description            | Possible Values                                                             |
| :----- | :------ | :----- | :------ | :--------------------- | :-------------------------------------------------------------------------- |
| `x`    | number  | inches | `1.0`   | horizontal location    | 0-n OR 'n%'. (Ex: `{x:'50%'}` will place object in the middle of the Slide) |
| `y`    | number  | inches | `1.0`   | vertical location      | 0-n OR 'n%'.                                                                |
| `w`    | number  | inches |         | width                  | 0-n OR 'n%'. (Ex: `{w:'50%'}` will make object 50% width of the Slide)      |
| `h`    | number  | inches |         | height                 | 0-n OR 'n%'.                                                                |
| `colW` | integer | inches |         | width for every column | Ex: Width for every column in table (uniform) `2.0`                         |
| `colW` | array   | inches |         | column widths in order | Ex: Width for each of 5 columns `[1.0, 2.0, 2.5, 1.5, 1.0]`                 |
| `rowH` | integer | inches |         | height for every row   | Ex: Height for every row in table (uniform) `2.0`                           |
| `rowH` | array   | inches |         | row heights in order   | Ex: Height for each of 5 rows `[1.0, 2.0, 2.5, 1.5, 1.0]`                   |

## Table Auto-Paging Options

| Option          | Type          | Default | Description                                 | Possible Values                           |
| :-------------- | :------------ | :------ | :------------------------------------------ | :---------------------------------------- |
| `autoPage`      | boolean       | `true`  | auto-page table                             | `true` or `false`. Ex: `{autoPage:false}` |
| `lineWeight`    | float         | 0       | line weight value                           | -1.0 to 1.0. Ex: `{lineWeight:0.5}`       |
| `newPageStartY` | number/string |         | starting `y` value for tables on new Slides | 0-n OR 'n%'. Ex:`{newPageStartY:0.5}`     |

### Option Details

- `autoPage`: allows the auto-paging functionality (as table rows overflow the Slide, new Slides will be added) to be disabled.
- `lineWeight`: adjusts the calculated height of lines. If too much empty space is left under each table,
  then increase lineWeight value. Conversely, if the tables are overflowing the bottom of the Slides, then
  reduce the lineWeight value. Also helpful when using some fonts that do not have the usual golden ratio.
- `newPageStartY`: provides the ability to specify where new tables will be placed on new Slides. For example,
  you may place a table halfway down a Slide, but you wouldn't that to be the starting location for subsequent
  tables. Use this option to ensure there is no wasted space and to guarantee a professional look.

## Table Auto-Paging Notes

- New slides will be created as tables overflow. The table will start at either `newPageStartY` (if present) or the Slide's top `margin`.
- Tables will retain their existing `x`, `w`, and `colW` values as they are rendered onto subsequent Slides.
- Auto-paging is not an exact science! Try using different `lineWeight` and Slide margin values if your tables are overflowing the Slide.
- There are many examples of auto-paging in the `examples` folder.

## Table Formatting Options

| Option      | Type    | Unit   | Default | Description        | Possible Values                                                                   |
| :---------- | :------ | :----- | :------ | :----------------- | :-------------------------------------------------------------------------------- |
| `align`     | string  |        | `left`  | alignment          | `left` or `center` or `right`                                                     |
| `bold`      | boolean |        | `false` | bold text          | `true` or `false`                                                                 |
| `border`    | object  |        |         | cell border        | object with `type`, `pt` and `color` values. (see below)                          |
| `border`    | array   |        |         | cell border        | array of objects with `pt` and `color` values in TRBL order.                      |
| `color`     | string  |        |         | text color         | hex color code or [scheme color constant](#scheme-colors). Ex: `{color:'0088CC'}` |
| `colspan`   | integer |        |         | column span        | 2-n. Ex: `{colspan:2}` (Note: be sure to include a table `w` value)               |
| `fill`      | string  |        |         | fill/bkgd color    | hex color code or [scheme color constant](#scheme-colors). Ex: `{color:'0088CC'}` |
| `fontFace`  | string  |        |         | font face          | Ex: `{fontFace:'Arial'}`                                                          |
| `fontSize`  | number  | points |         | font size          | 1-256. Ex: `{fontSize:12}`                                                        |
| `italic`    | boolean |        | `false` | italic text        | `true` or `false`                                                                 |
| `margin`    | number  | points |         | margin             | 0-99 (ProTip: use the same value from CSS `padding`)                              |
| `margin`    | array   | points |         | margin             | array of integer values in TRBL order. Ex: `margin:[5,10,5,10]`                   |
| `rowspan`   | integer |        |         | row span           | 2-n. Ex: `{rowspan:2}`                                                            |
| `underline` | boolean |        | `false` | underline text     | `true` or `false`                                                                 |
| `valign`    | string  |        |         | vertical alignment | `top` or `middle` or `bottom` (or `t` `m` `b`)                                    |

### Border Option
| Option      | Type    | Default | Description        | Possible Values                                                                   |
| :---------- | :------ | :------ | :----------------- | :-------------------------------------------------------------------------------- |
| `type`      | string  | `solid` | border type        | `none` or `solid` or `dash`                                                       |
| `pt`        | string  | `1`     | border thickness   | any positive number                                                               |
| `color`     | string  | `black` | cell border        | hex color code or [scheme color constant](#scheme-colors). Ex: `{color:'0088CC'}` |

## Table Formatting Notes

- **Formatting Options** passed to `slide.addTable()` apply to every cell in the table
- You can selectively override formatting at a cell-level providing any **Formatting Option** in the cell `options`

## Table Cell Formatting

- Table cells can be either a plain text string or an object with text and options properties
- When using an object, any of the formatting options above can be passed in `options` and will apply to that cell only
- Cell borders can be removed (aka: borderless table) by using the 'none' type (Ex: `border: {type:'none'}`)

Bullets and word-level formatting are supported inside table cells. Passing an array of objects with text/options values
as the `text` value allows fine-grained control over the text inside cells.

- Available formatting options are here: [Text Options](/PptxGenJS/docs/api-text.html#text-options)
- See below for examples or view the `demos/browser/index.html` page for lots more

## Table Cell Formatting Examples

```javascript
// TABLE 1: Cell-level Formatting
var rows = [];
// Row One: cells will be formatted according to any options provided to `addTable()`
rows.push(["First", "Second", "Third"]);
// Row Two: set/override formatting for each cell
rows.push([
  { text: "1st", options: { color: "ff0000" } },
  { text: "2nd", options: { color: "00ff00" } },
  { text: "3rd", options: { color: "0000ff" } }
]);
slide.addTable(rows, { x: 0.5, y: 1.0, w: 9.0, color: "363636" });

// TABLE 2: Using word-level formatting inside cells
// NOTE: An array of text/options objects provides fine-grained control over formatting
var arrObjText = [
  { text: "Red ", options: { color: "FF0000" } },
  { text: "Green ", options: { color: "00FF00" } },
  { text: "Blue", options: { color: "0000FF" } }
];
// EX A: Pass an array of text objects to `addText()`
slide.addText(arrObjText, {
  x: 0.5,
  y: 2.0,
  w: 9,
  h: 1,
  margin: 0.1,
  fill: "232323"
});

// EX B: Pass the same objects as a cell's `text` value
var arrTabRows = [
  [
    { text: "Cell 1 A", options: { fontFace: "Arial" } },
    { text: "Cell 1 B", options: { fontFace: "Courier" } },
    { text: arrObjText, options: { fill: "232323" } }
  ]
];
slide.addTable(arrTabRows, { x: 0.5, y: 3.5, w: 9, h: 1, colW: [1.5, 1.5, 6] });
```

## Table Examples

```javascript
var pptx = new PptxGenJS();
var slide = pptx.addSlide();
slide.addText("Demo-03: Table", {
  x: 0.5,
  y: 0.25,
  fontSize: 18,
  fontFace: "Arial",
  color: "0088CC"
});

// TABLE 1: Single-row table
// --------
var rows = [["Cell 1", "Cell 2", "Cell 3"]];
var tabOpts = {
  x: 0.5,
  y: 1.0,
  w: 9.0,
  fill: "F7F7F7",
  fontSize: 14,
  color: "363636"
};
slide.addTable(rows, tabOpts);

// TABLE 2: Multi-row table (each rows array element is an array of cells)
// --------
var rows = [["A1", "B1", "C1"], ["A2", "B2", "C2"]];
var tabOpts = {
  x: 0.5,
  y: 2.0,
  w: 9.0,
  fill: "F7F7F7",
  fontSize: 18,
  color: "6f9fc9"
};
slide.addTable(rows, tabOpts);

// TABLE 3: Formatting at a cell level - use this to selectively override table's cell options
// --------
var rows = [
  [
    {
      text: "Top Lft",
      options: { valign: "top", align: "left", fontFace: "Arial" }
    },
    {
      text: "Top Ctr",
      options: { valign: "top", align: "center", fontFace: "Verdana" }
    },
    {
      text: "Top Rgt",
      options: { valign: "top", align: "right", fontFace: "Courier" }
    }
  ]
];
var tabOpts = {
  x: 0.5,
  y: 4.5,
  w: 9.0,
  rowH: 0.6,
  fill: "F7F7F7",
  fontSize: 18,
  color: "6f9fc9",
  valign: "center"
};
slide.addTable(rows, tabOpts);

// Multiline Text / Line Breaks - use either "\r" or "\n"
slide.addTable([["Line 1\nLine 2\nLine 3"]], { x: 2, y: 3, w: 4 });

pptx.writeFile("Demo-Tables");
```
