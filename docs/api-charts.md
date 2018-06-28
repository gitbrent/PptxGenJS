---
id: api-charts
title: Adding Charts
---
## Syntax
```javascript
slide.addChart({TYPE}, {DATA}, {OPTIONS});
```

## Chart Types
* Chart type can be any one of `pptx.charts`
* Currently: `pptx.charts.AREA`, `pptx.charts.BAR`, `pptx.charts.BUBBLE`, `pptx.charts.LINE`, `pptx.charts.SCATTER`, `pptx.charts.PIE`, `pptx.charts.DOUGHNUT`

## Multi-Type Charts
* Chart types can be any one of `pptx.charts`, although `pptx.charts.AREA`, `pptx.charts.BAR`, and `pptx.charts.LINE` will give the best results.
* There should be at least two chart-types. There should always be two value axes and category axes.
* Multi Charts have a different function signature than standard. There are two parameters:
 * `chartTypes`: Array of objects, each with `type`, `data`, and `options` objects.
 * `options`: Standard options as used with single charts. Can include axes options.
* Columns makes the most sense in general. Line charts cannot be rotated to match up with horizontal bars (a PowerPoint limitation).
* Can optionally have a secondary value axis.
* If there is secondary value axis, a secondary category axis is required in order to render, but currently always uses the primary labels. It is recommended to use `catAxisHidden: true` on the secondary category axis.
* Standard options are used, and the chart-type-options are mixed in to each.

## Multi-Type Syntax
```javascript
slide.addChart({MULTI_TYPES_AND_DATA}, {OPTIONS_AND_AXES});
```

## Chart Notes
* Zero values can be hidden using Microsoft formatting specs (see [Issue #288](https://github.com/gitbrent/PptxGenJS/issues/278))

## Chart Size/Formatting Options
| Option          | Type    | Unit    | Default   | Description                        | Possible Values  |
| :-------------- | :------ | :------ | :-------- | :--------------------------------- | :--------------- |
| `x`             | number  | inches  | `1.0`     | horizontal location                | 0-n OR 'n%'. (Ex: `{x:'50%'}` places object in middle of the Slide) |
| `y`             | number  | inches  | `1.0`     | vertical location                  | 0-n OR 'n%'. |
| `w`             | number  | inches  | `50%`     | width                              | 0-n OR 'n%'. (Ex: `{w:'50%'}` will make object 50% width of the Slide) |
| `h`             | number  | inches  | `50%`     | height                             | 0-n OR 'n%'. |
| `border`        | object  |         |           | chart border                  | object with `pt` and `color` values. Ex: `border:{pt:'1', color:'f1f1f1'}` |
| `chartColors`   | array   |         |           | data colors                        | array of hex color codes. Ex: `['0088CC','FFCC00']` |
| `chartColorsOpacity` | number | percent | `100` | data color opacity percent         | 1-100. Ex: `{ chartColorsOpacity:50 }` |
| `fill`          | string  |         |           | fill/background color              | hex color code. Ex: `{ fill:'0088CC' }` |
| `holeSize`      | number  | percent | `50`      | doughnut hole size                 | 1-100. Ex: `{ holeSize:50 }` |
| `invertedColors`| array   |         |           | data colors for negative numbers   | array of hex color codes. Ex: `['0088CC','FFCC00']` |
| `legendFontSize`| number  | points  | `10`      | legend font size                   | 1-256. Ex: `{ legendFontSize: 13 }`|
| `legendColor`   | string  |         | `000000`  | legend text color                  | hex color code. Ex: `{ legendColor: '0088CC' }` |
| `legendPos`     | string  |         | `r`       | chart legend position              | `b` (bottom), `tr` (top-right), `l` (left), `r` (right), `t` (top) |
| `layout`        | object  |         |           | positioning plot within chart area | object with `x`, `y`, `w` and `h` props, all in range 0-1 (proportionally related to the chart size). Ex: `{x: 0, y: 0, w: 1, h: 1}` fully expands chart within the plot area |
| `showDataTable`           | boolean |     | `false`   | show Data Table under the chart     | `true` or `false` (Not available for Pie/Doughnut charts) |
| `showDataTableKeys`       | boolean |     | `true`    | show Data Table Keys (color blocks) | `true` or `false` (Not available for Pie/Doughnut charts) |
| `showDataTableHorzBorder` | boolean |     | `true`    | show Data Table horizontal borders  | `true` or `false` (Not available for Pie/Doughnut charts) |
| `showDataTableVertBorder` | boolean |     | `true`    | show Data Table vertical borders    | `true` or `false` (Not available for Pie/Doughnut charts) |
| `showDataTableOutline`    | boolean |     | `true`    | show Data Table table outline       | `true` or `false` (Not available for Pie/Doughnut charts) |
| `showLabel`     | boolean |         | `false`   | show data labels                   | `true` or `false` |
| `showLegend`    | boolean |         | `false`   | show chart legend                  | `true` or `false` |
| `showPercent`   | boolean |         | `false`   | show data percent                  | `true` or `false` |
| `showTitle`     | boolean |         | `false`   | show chart title                   | `true` or `false` |
| `showValue`     | boolean |         | `false`   | show data values                   | `true` or `false` |
| `title`         | string  |         |           | chart title                        | a string. Ex: `{ title:'Sales by Region' }` |
| `titleAlign`    | string  |         | `center`  | chart title text align             | `left` `center` or `right` Ex: `{ titleAlign:'left' }` |
| `titleColor`    | string  |         | `000000`  | title color                        | hex color code. Ex: `{ titleColor:'0088CC' }` |
| `titleFontFace` | string  |         | `Arial`   | font face                          | font name. Ex: `{ titleFontFace:'Arial' }` |
| `titleFontSize` | number  | points  | `18`      | font size                          | 1-256. Ex: `{ titleFontSize:12 }` |
| `titlePos`      | object  |         |           | title position                     | object with x and y values. Ex: `{ titlePos:{x: 0, y: 10} }` |
| `titleRotate`   | integer | degrees |           | title rotation degrees             | 0-360. Ex: `{ titleRotate:45 }` |

## Chart Axis Options
| Option                 | Type    | Unit    | Default      | Description                   | Possible Values                                  |
| :--------------------- | :------ | :------ | :----------- | :---------------------------- | :----------------------------------------------- |
| `axisLineColor`        | string  |         | `000000`     | cat/val axis line color       | hex color code. Ex: `{ axisLineColor:'0088CC' }` |
| `catAxisBaseTimeUnit`  | string  |         |              | category-axis base time unit  | `days` `months` or `years` |
| `catAxisHidden`        | boolean |         | `false`      | hide category-axis            | `true` or `false`   |
| `catAxisLabelColor`    | string  |         | `000000`     | category-axis color           | hex color code. Ex: `{ catAxisLabelColor:'0088CC' }`   |
| `catAxisLabelFontFace` | string  |         | `Arial`      | category-axis font face       | font name. Ex: `{ titleFontFace:'Arial' }` |
| `catAxisLabelFontSize` | integer | points  | `18`         | category-axis font size       | 1-256. Ex: `{ titleFontSize:12 }`          |
| `catAxisLabelFrequency`| integer |         |              | PPT "Interval Between Labels" | 1-n. Ex: `{ catAxisLabelFrequency: 2 }`          |
| `catAxisLabelPos`      | string  | string  | `nextTo`     | axis label position     | `low`, `high`, or `nextTo` . Ex: `{ catAxisLabelPos: 'low' }`      |
| `catAxisLineShow`      | boolean |         | true         | show/hide category-axis line  | `true` or `false` |
| `catAxisMajorTimeUnit` | string  |         |              | category-axis major time unit | `days` `months` or `years` |
| `catAxisMinorTimeUnit` | string  |         |              | category-axis minor time unit | `days` `months` or `years` |
| `catAxisMajorUnit`     | integer |         |              | category-axis major unit      | Positive integer. Ex: `{ catAxisMajorUnit:12 }`   |
| `catAxisMinorUnit`     | integer |         |              | category-axis minor unit      | Positive integer. Ex: `{ catAxisMinorUnit:1 }`   |
| `catAxisOrientation`   | string  |         | `minMax`     | category-axis orientation     | `maxMin` (high->low) or `minMax` (low->high) |
| `catAxisTitle`         | string  |         | `Axis Title` | axis title                    | a string. Ex: `{ catAxisTitle:'Regions' }` |
| `catAxisTitleColor`    | string  |         | `000000`     | title color                   | hex color code. Ex: `{ catAxisTitleColor:'0088CC' }` |
| `catAxisTitleFontFace` | string  |         | `Arial`      | font face                     | font name. Ex: `{ catAxisTitleFontFace:'Arial' }` |
| `catAxisTitleFontSize` | integer | points  |              | font size                     | 1-256. Ex: `{ catAxisTitleFontSize:12 }` |
| `catAxisTitleRotate`   | integer | degrees |              | title rotation degrees        | 0-360. Ex: `{ catAxisTitleRotate:45 }` |
| `catGridLine`          | object  |         | `none`       | category grid line style      | object with properties `size` (pt), `color` and `style` (`'solid'`, `'dash'` or `'dot'`) or `'none'` to hide |
| `showCatAxisTitle`     | boolean |         | `false`      | show category (vert) title   | `true` or `false`.  Ex:`{ showCatAxisTitle:true }` |
| `showValAxisTitle`     | boolean |         | `false`      | show values (horiz) title    | `true` or `false`.  Ex:`{ showValAxisTitle:true }` |
| `valAxisHidden`        | boolean |         | `false`      | hide value-axis              | `true` or `false`   |
| `valAxisLabelColor`    | string  |         | `000000`     | value-axis color             | hex color code. Ex: `{ valAxisLabelColor:'0088CC' }` |
| `valAxisLabelFontFace` | string  |         | `Arial`      | value-axis font face         | font name. Ex: `{ titleFontFace:'Arial' }`   |
| `valAxisLabelFontSize` | integer | points  | `18`         | value-axis font size         | 1-256. Ex: `{ titleFontSize:12 }`            |
| `valAxisLabelFormatCode` | string |        | `General`    | value-axis number format     | format string. Ex: `{ axisLabelFormatCode:'#,##0' }` [MicroSoft Number Format Codes](https://support.office.com/en-us/article/Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68) |
| `valAxisLineShow`      | boolean |         | true         | show/hide value-axis line    | `true` or `false` |
| `valAxisMajorUnit`     | number  |         | `1.0`        | value-axis tick steps        | Float or whole number. Ex: `{ majorUnit:0.2 }`      |
| `valAxisMaxVal`        | number  |         |              | value-axis maximum value     | 1-N. Ex: `{ valAxisMaxVal:125 }` |
| `valAxisMinVal`        | number  |         |              | value-axis minimum value     | 1-N. Ex: `{ valAxisMinVal: -10 }` |
| `valAxisOrientation`   | string  |         | `minMax`     | value-axis orientation       | `maxMin` (high->low) or `minMax` (low->high) |
| `valAxisTitle`         | string  |         | `Axis Title` | axis title                   | a string. Ex: `{ valAxisTitle:'Sales (USD)' }` |
| `valAxisTitleColor`    | string  |         | `000000`     | title color                  | hex color code. Ex: `{ valAxisTitleColor:'0088CC' }` |
| `valAxisTitleFontFace` | string  |         | `Arial`      | font face                    | font name. Ex: `{ valAxisTitleFontFace:'Arial' }` |
| `valAxisTitleFontSize` | number  | points  |              | font size                    | 1-256. Ex: `{ valAxisTitleFontSize:12 }` |
| `valAxisTitleRotate`   | integer | degrees |              | title rotation degrees       | 0-360. Ex: `{ valAxisTitleRotate:45 }` |
| `valGridLine`          | object  |         |              | value grid line style        | object with properties `size` (pt), `color` and `style` (`'solid'`, `'dash'` or `'dot'`) or `'none'` to hide |

## Chart Data Options
| Option                 | Type    | Unit    | Default   | Description                | Possible Values                            |
| :--------------------- | :------ | :------ | :-------- | :------------------------- | :----------------------------------------- |
| `barDir`               | string  |         | `col`     | bar direction        | (*Bar Chart*) `bar` (horizontal) or `col` (vertical). Ex: `{barDir:'bar'}` |
| `barGapWidthPct`       | number  | percent | `150`     | width % between bar groups | (*Bar Chart*) 0-999. Ex: `{ barGapWidthPct:50 }` |
| `barGrouping`          | string  |         |`clustered`| bar grouping               | (*Bar Chart*) `clustered` or `stacked` or `percentStacked`. |
| `dataBorder`           | object  |         |           | data border          | object with `pt` and `color` values. Ex: `border:{pt:'1', color:'f1f1f1'}` |
| `dataLabelColor`       | string  |         | `000000`  | data label color           | hex color code. Ex: `{ dataLabelColor:'0088CC' }`     |
| `dataLabelFormatCode`  | string  |         |           | format to show data value  | format string. Ex: `{ dataLabelFormatCode:'#,##0' }` [MicroSoft Number Format Codes](https://support.office.com/en-us/article/Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68)  |
| `dataLabelFontFace`    | string  |         | `Arial`   | value-axis font face       | font name. Ex: `{ titleFontFace:'Arial' }`   |
| `dataLabelFontSize`    | number  | points  | `18`      | value-axis font size       | 1-256. Ex: `{ titleFontSize:12 }`            |
| `dataLabelPosition`    | string  |         | `bestFit` | data label position        | `bestFit`,`b`,`ctr`,`inBase`,`inEnd`,`l`,`outEnd`,`r`,`t` |
| `dataNoEffects`        | boolean |         | `false`   | whether to omit effects on data | (*Doughnut/Pie Charts*) `true` or `false` |
| `displayBlanksAs`      | string  |         | `span`    | whether to draw line or gap | (*Line Charts*) `span` or `gap` |
| `gridLineColor`        | string  |         | `000000`  | grid line color            | hex color code. Ex: `{ gridLineColor:'0088CC' }`     |
| `lineDataSymbol`       | string  |         | `circle`  | symbol used on line marker | `circle`,`dash`,`diamond`,`dot`,`none`,`square`,`triangle` |
| `lineDataSymbolSize`   | number  | points  | `6`       | size of line data symbol   | 1-256. Ex: `{ lineDataSymbolSize:12 }` |
| `lineDataSymbolLineSize` | number | points | `0.75`    | size of data symbol outline   | 1-256. Ex: `{ lineDataSymbolLineSize:12 }` |
| `lineDataSymbolLineColor`| number | points | `0.75`    | size of data symbol outline   | 1-256. Ex: `{ lineDataSymbolLineSize:12 }` |
| `lineSize`             | number  | points  | `2`       | thickness of data line (0 is no line) | 0-256. Ex: `{ lineSize: 1 }` |
| `lineSmooth`           | boolean |         | `false`   | whether to smooth lines | `true` or `false` | Ex: `{ lineSmooth: true }` |
| `shadow`               | object  |         |           | data element shadow options   | `'none'` or [shadow options](#chart-element-shadow-options) |
| `valueBarColors`       | boolean |         | `false`   | forces chartColors on multi-data-series | `true` or `false` |

## Chart Element Shadow Options
| Option       | Type    | Unit    | Default   | Description         | Possible Values                            |
| :----------- | :------ | :------ | :-------- | :------------------ | :----------------------------------------- |
| `type`       | string  |         | `outer`   | shadow type         | `outer` or `inner`. Ex: `{ type:'outer' }` |
| `angle`      | number  | degrees | `90`      | shadow angle        | 0-359. Ex: `{ angle:90 }`                  |
| `blur`       | number  | points  | `3`       | blur size           | 1-256. Ex: `{ blur:3 }`                    |
| `color`      | string  |         | `000000`  | shadow color        | hex color code. Ex: `{ color:'0088CC' }`   |
| `offset`     | number  | points  | `1.8`     | offset size         | 1-256. Ex: `{ offset:2 }`                  |
| `opacity`    | number  | percent | `0.35`    | opacity             | 0-1. Ex: `{ opacity:0.35 }`                |

## Chart Multi-Type Options
| Option             | Type    | Default  | Description                                             | Possible Values   |
| :----------------- | :------ | :------- | :------------------------------------------------------ | :---------------- |
| `catAxes`          | array   |          | array of two axis options objects | See example below   |                   |
| `secondaryCatAxis` | boolean | `false`  | If data should use secondary category axis (or primary) | `true` or `false` |
| `secondaryValAxis` | boolean | `false`  | If data should use secondary value axis (or primary)    | `true` or `false` |
| `valAxes`          | array   |          | array of two axis options objects | See example below   |                   |

## Chart Examples
```javascript
var pptx = new PptxGenJS();
pptx.setLayout('LAYOUT_WIDE');

var slide = pptx.addNewSlide();

// Chart Type: BAR
var dataChartBar = [
  {
    name  : 'Region 1',
    labels: ['May', 'June', 'July', 'August'],
    values: [26, 53, 100, 75]
  },
  {
    name  : 'Region 2',
    labels: ['May', 'June', 'July', 'August'],
    values: [43.5, 70.3, 90.1, 80.05]
  }
];
slide.addChart( pptx.charts.BAR, dataChartBar, { x:1.0, y:1.0, w:12, h:6 } );

// Chart Type: AREA
// Chart Type: LINE
var dataChartAreaLine = [
  {
    name  : 'Actual Sales',
    labels: ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
    values: [1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123, 15121]
  },
  {
    name  : 'Projected Sales',
    labels: ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],
    values: [1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123, 12121]
  }
];
slide.addChart( pptx.charts.AREA, dataChartAreaLine, { x:1.0, y:1.0, w:12, h:6 } );
slide.addChart( pptx.charts.LINE, dataChartAreaLine, { x:1.0, y:1.0, w:12, h:6 } );

// Chart Type: PIE
var dataChartPie = [
  { name:'Location', labels:['DE','GB','MX','JP','IN','US'], values:[35,40,85,88,99,101] }
];
slide.addChart( pptx.charts.PIE, dataChartPie, { x:1.0, y:1.0, w:6, h:6 } );

// Chart Type: XY SCATTER
var dataChartScatter = [
	{ name:'X-Axis',    values:[1,2,3,4,5,6,7,8,9,10] },
	{ name:'Y-Value 1', values:[13, 20, 21, 25] },
	{ name:'Y-Value 2', values:[21, 22, 25, 49] }
];
slide.addChart( pptx.charts.SCATTER, dataChartScatter, { x:1.0, y:1.0, w:6, h:4 } );

// Chart Type: BUBBLE
var dataChartBubble = [
	{ name:'X-Axis',   values:[1, 2, 3, 4, 5, 6] },
	{ name:'Airplane', values:[33, 20, 51, 65, 71, 75], sizes:[10,10,12,12,15,20] },
	{ name:'Train',    values:[99, 88, 77, 89, 99, 99], sizes:[20,20,22,22,25,30] },
	{ name:'Bus',      values:[21, 22, 25, 49, 59, 69], sizes:[11,11,13,13,16,21] }
];
slide.addChart( pptx.charts.BUBBLE, dataChartBubble, { x:1.0, y:1.0, w:6, h:4 } );

// Chart Type: Multi-Type
// NOTE: use the same labels for all types
var labels = ['Q1', 'Q2', 'Q3', 'Q4', 'OT'];
var chartTypes = [
  {
    type: pptx.charts.BAR,
    data: [{
      name: 'Projected',
      labels: labels,
      values: [17, 26, 53, 10, 4]
    }],
    options: { barDir: 'col' }
  },
  {
    type: pptx.charts.LINE,
    data: [{
      name: 'Current',
      labels: labels,
      values: [5, 3, 2, 4, 7]
    }],
    options: {
      // NOTE: both are required, when using a secondary axis:
      secondaryValAxis: true,
      secondaryCatAxis: true
    }
  }
];
var multiOpts = {
  x:1.0, y:1.0, w:6, h:6,
  showLegend: false,
  valAxisMaxVal: 100,
  valAxisMinVal: 0,
  valAxisMajorUnit: 20,
  valAxes:[
    {
      showValAxisTitle: true,
      valAxisTitle: 'Primary Value Axis'
    },
    {
      showValAxisTitle: true,
      valAxisTitle: 'Secondary Value Axis',
      valAxisMajorUnit: 1,
      valAxisMaxVal: 10,
      valAxisMinVal: 1,
      valGridLine: 'none'
    }
  ],
  catAxes: [{ catAxisTitle: 'Primary Category Axis' }, { catAxisHidden: true }]
};

slide.addChart(chartTypes, multiOpts);

pptx.save('Demo-Chart');
```
