# Change Log

## [v1.9.0](https://github.com/gitbrent/pptxgenjs/tree/v1.8.0) (2017-10-10)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.8.0...v1.9.0)

**Fixed Bugs:**
- Vertical align and line break bug since update [\#79](https://github.com/gitbrent/PptxGenJS/issues/79) ([mirkoint](https://github.com/mirkoint))
- Save callback is not called by client-browser when there are images to encode [\#187](https://github.com/gitbrent/PptxGenJS/issues/187) ([Malangs](https://github.com/Malangs))
- Promise Dependency - TypeError: Promise.all is not a function [\#188](https://github.com/gitbrent/PptxGenJS/issues/188) ([bartolomeu](https://github.com/bartolomeu))
- Default text size in empty cells making row height too big [\#193](https://github.com/gitbrent/PptxGenJS/issues/193) ([mreilaender](https://github.com/mreilaender))
- Fixed issue that included many extraneous tab characters in the table demo lorem-ipsum text (GitBrent)
- Fix chart issue: Entities encoding [\#204](https://github.com/gitbrent/PptxGenJS/pull/204) ([clubajax](https://github.com/clubajax))
- Fix chart issue: val axis [\#205](https://github.com/gitbrent/PptxGenJS/pull/205) ([clubajax](https://github.com/clubajax))
- Fix chart issue: Line chart series colors were not being respected [\#206](https://github.com/gitbrent/PptxGenJS/pull/206) ([kyrrigle](https://github.com/kyrrigle))
- Discrepancy between docs and code regarding setting a slide's background [\#207](https://github.com/gitbrent/PptxGenJS/pull/207) ([msambarino](https://github.com/msambarino))
- Fix chart issue: bar color regression [\#210](https://github.com/gitbrent/PptxGenJS/pull/210) ([clubajax](https://github.com/clubajax))

**Implemented Enhancements:**
- New chart feature: category axis dates [\#149](https://github.com/gitbrent/PptxGenJS/pull/149) ([kyrrigle](https://github.com/kyrrigle))
- New image option: sizing [\#177](https://github.com/gitbrent/PptxGenJS/pull/177) ([kajda90](https://github.com/kajda90))
- New chart option: show Data Table [\#182](https://github.com/gitbrent/PptxGenJS/issues/182) ([akashkarpe](https://github.com/akashkarpe))
- New chart option: catAxisLabelFrequency [\#184](https://github.com/gitbrent/PptxGenJS/pull/184) ([kajda90](https://github.com/kajda90))
- New chart type: XY Scatter [\#192](https://github.com/gitbrent/PptxGenJS/issues/192) ([shaunvdp](https://github.com/shaunvdp))
- Add electron detection to load correct jquery version [\#200](https://github.com/gitbrent/PptxGenJS/issues/200) ([mreilaender](https://github.com/mreilaender))



## [v1.8.0](https://github.com/gitbrent/pptxgenjs/tree/v1.8.0) (2017-09-12)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.7.0...v1.8.0)

**Fixed Bugs:**
- Slide numbers wrap over 99 [\#133](https://github.com/gitbrent/PptxGenJS/issues/133) ([sangramjagtap](https://github.com/sangramjagtap))
- Shadow corrections bugfix [\#136](https://github.com/gitbrent/PptxGenJS/pull/136) ([kajda90](https://github.com/kajda90))
- Negative Chart values throwing error [\#175](https://github.com/gitbrent/PptxGenJS/issues/175) ([shaunvdp](https://github.com/shaunvdp))

**Implemented Enhancements:**
- New chart feature: Bar colors and axis [\#132](https://github.com/gitbrent/PptxGenJS/pull/132) ([clubajax](https://github.com/clubajax))
- New feature: Scheme colors [\#135](https://github.com/gitbrent/PptxGenJS/pull/135) ([kajda90](https://github.com/kajda90))
- New chart feature: lineShadow [\#138](https://github.com/gitbrent/PptxGenJS/pull/138) ([kajda90](https://github.com/kajda90))
- New chart type: Tornado Chart [\#140](https://github.com/gitbrent/PptxGenJS/pull/140) ([clubajax](https://github.com/clubajax))
- New chart feature: layout option [\#141](https://github.com/gitbrent/PptxGenJS/pull/141) ([kajda90](https://github.com/kajda90))
- New chart type: Doughnut Chart [\#142](https://github.com/gitbrent/PptxGenJS/pull/142) ([kyrrigle](https://github.com/kyrrigle))
- New chart options: gridlines and axes [\#143](https://github.com/gitbrent/PptxGenJS/pull/143) ([kajda90](https://github.com/kajda90))
- New chart feature: Axis Titles [\#144](https://github.com/gitbrent/PptxGenJS/pull/144) ([kyrrigle](https://github.com/kyrrigle))
- Optional output type [\#147](https://github.com/gitbrent/PptxGenJS/pull/147) ([kajda90](https://github.com/kajda90))
- New chart options: catAxisLineShow [\#152](https://github.com/gitbrent/PptxGenJS/pull/152) ([amgault](https://github.com/amga))
- New Master Slide Layouts [\#161](https://github.com/gitbrent/PptxGenJS/pull/161) ([kajda90](https://github.com/kajda90))
- Demo page updates [\#164](https://github.com/gitbrent/PptxGenJS/pull/164) ([clubajax](https://github.com/clubajax))
- New chart feature: New Legend/Title Options [\#165](https://github.com/gitbrent/PptxGenJS/pull/165) ([clubajax](https://github.com/clubajax))
- New chart options: Shadows and Transparent Color [\#166](https://github.com/gitbrent/PptxGenJS/pull/166) ([clubajax](https://github.com/clubajax))
- Add no border option to tables [\#169](https://github.com/gitbrent/PptxGenJS/issues/169) ([eddyclock](https://github.com/eddyclock))
- Chart: Escape Labels XML [\#171](https://github.com/gitbrent/PptxGenJS/pull/171) ([kyrrigle](https://github.com/kyrrigle))
- Add new 'lang' text option to enable Chinese Word fonts [\#174](https://github.com/gitbrent/PptxGenJS/issues/174) ([eddyclock](https://github.com/eddyclock))
- Add color validation to createColorElement() [\#178](https://github.com/gitbrent/PptxGenJS/pull/178) ([kajda90](https://github.com/kajda90))



## [v1.7.0](https://github.com/gitbrent/pptxgenjs/tree/v1.7.0) (2017-08-07)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.6.0...v1.7.0)

**Fixed Bugs:**
- Unable to edit data on line chart [\#122](https://github.com/gitbrent/PptxGenJS/issues/122) ([david23zhu](https://github.com/david23zhu))

**Implemented Enhancements:**
- Add charts to Masters/Templates [\#114](https://github.com/gitbrent/PptxGenJS/issues/114) ([yipiha](https://github.com/yipiha))
- Format text as a superscript in a table cell [\#120](https://github.com/gitbrent/PptxGenJS/issues/120) ([aranard](https://github.com/aranard))



## [v1.6.0](https://github.com/gitbrent/pptxgenjs/tree/v1.6.0) (2017-07-17)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.5.0...v1.6.0)

**Fixed Bugs:**
- The width or the height must be an integer not a float [\#29](https://github.com/gitbrent/PptxGenJS/issues/29) ([badlee](https://github.com/badlee))

**Implemented Enhancements:**
- HTTP Stream [\#35](https://github.com/gitbrent/PptxGenJS/issues/35) ([FedeMM](https://github.com/FedeMM))
- Add a 'line spacing' option to addText() [\#104](https://github.com/gitbrent/PptxGenJS/issues/104) ([eddyclock](https://github.com/eddyclock))
- err TypeError: Cannot read property 'text' of undefined [\#106](https://github.com/gitbrent/PptxGenJS/issues/106) ([ninas880025](https://github.com/ninas880025))
- Added bowser support, gulp build of bundle [\#107](https://github.com/gitbrent/PptxGenJS/pull/107) ([santi-git](https://github.com/santi-git))
- Add increase/decrease indent for bullets [\#108](https://github.com/gitbrent/PptxGenJS/issues/108) ([sangramjagtap](https://github.com/sangramjagtap))



## [v1.5.0](https://github.com/gitbrent/pptxgenjs/tree/v1.5.0) (2017-05-26)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.4.0...v1.5.0)

**Fixed Bugs:**
- Hyperlink and font_face problem [\#74](https://github.com/gitbrent/PptxGenJS/issues/74) ([ZouhaierSebri](https://github.com/ZouhaierSebri))
- Can't override margin with 0 [\#78](https://github.com/gitbrent/PptxGenJS/issues/78) ([scottmtraver](https://github.com/scottmtraver))
- Issue with autopage and colspan [\#80](https://github.com/gitbrent/PptxGenJS/issues/80) ([Szymon-dziewonski](https://github.com/Szymon-dziewonski))
- Does not work on Firefox for Mac, no issues on Firefox for windows [\#81](https://github.com/gitbrent/PptxGenJS/issues/81) ([alexanderdevm](https://github.com/alexanderdevm) and [rwhitmore90](https://github.com/rwhitmore90))
- Not a real issue, just a quick README fix [\#88](https://github.com/gitbrent/PptxGenJS/issues/88) ([mirkoint](https://github.com/mirkoint))
- Invalid XML when calling .addText() with empty array [\#89](https://github.com/gitbrent/PptxGenJS/issues/89) ([JimmyTheChimp](https://github.com/JimmyTheChimp))
- Hyperlink and XML entities issue [\#90](https://github.com/gitbrent/PptxGenJS/issues/90) ([ZouhaierSebri](https://github.com/ZouhaierSebri))
- Tooltip option not implemented for image hyperlink [\#91](https://github.com/gitbrent/PptxGenJS/issues/91) ([ZouhaierSebri](https://github.com/ZouhaierSebri))

**Implemented Enhancements:**
- Add ability to create charts [\#51](https://github.com/gitbrent/PptxGenJS/issues/51) ([alagarrk](https://github.com/alagarrk))
- Added image type to shapes to allow images to be placed on top of shapes, added more properties to ppt document [\#53](https://github.com/gitbrent/PptxGenJS/pull/53) ([ericwgreene](https://github.com/ericwgreene))
- Add support for RTL (Right-to-Left) text for Arabic etc. [\#73](https://github.com/gitbrent/PptxGenJS/issues/73) ([vanekar](https://github.com/vanekar))
- Shape line Diagonal [\#75](https://github.com/gitbrent/PptxGenJS/issues/75) ([vanekar](https://github.com/vanekar))
- Add hyperlink to Image [\#77](https://github.com/gitbrent/PptxGenJS/issues/77) ([plopez7](https://github.com/plopez7))
- Adding rounding radius for texts and shapes and dash options for the outline [\#86](https://github.com/gitbrent/PptxGenJS/pull/86) ([ivolazy](https://github.com/ivolazy))



## [v1.4.0](https://github.com/gitbrent/pptxgenjs/tree/v1.4.0) (2017-04-10)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.3.0...v1.4.0)

**Fixed Bugs:**
- Auto Paging does not include master template on additional slides [\#61](https://github.com/gitbrent/PptxGenJS/issues/61) ([tb23911](https://github.com/tb23911))
- Issue calculating the available height for a table using Auto paging [\#64](https://github.com/gitbrent/PptxGenJS/issues/64) ([tb23911](https://github.com/tb23911))
- Multiple a:bodyPr tags within a:txBody causes damaged presentation in PowerPoint 2007 [\#69](https://github.com/gitbrent/PptxGenJS/issues/69) ([ZouhaierSebri](https://github.com/ZouhaierSebri))
- Text bug [\#71](https://github.com/gitbrent/PptxGenJS/issues/71) ([alexbai31](https://github.com/alexbai31))
- Errors when using Webpack/Typescript [\#72](https://github.com/gitbrent/PptxGenJS/issues/72) ([Vivihung](https://github.com/Vivihung))

**Implemented Enhancements:**
- Add Slide Number formatting options [\#68](https://github.com/gitbrent/PptxGenJS/issues/68) ([ZouhaierSebri](https://github.com/ZouhaierSebri))
- Added new feature: Hyperlinks as a text option

## [v1.3.0](https://github.com/gitbrent/pptxgenjs/tree/v1.3.0) (2017-03-22)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.2.1...v1.3.0)

**Fixed Bugs:**
- Added image type to shapes to allow images to be placed on top of shapes, added more properties to ppt document [\#53](https://github.com/gitbrent/PptxGenJS/pull/53) ([ericwgreene](https://github.com/ericwgreene))
- Table-to-Slides default for un-styled tables is black text on black bkgd [\#57](https://github.com/gitbrent/PptxGenJS/issues/57) ([orpitadutta](https://github.com/orpitadutta))
- Table Header and Auto Paging [\#62](https://github.com/gitbrent/PptxGenJS/issues/62) ([tb23911](https://github.com/tb23911))

**Implemented Enhancements:**
- Removed `FileSaver.js` as a required library (only JSZip and jQuery are required now)
- Allow text multi-formatting in single table cells [\#24](https://github.com/gitbrent/PptxGenJS/issues/24) ([jenkinsns](https://github.com/jenkinsns))
- Set fixed width to column using `addSlidesForTable()` [\#42](https://github.com/gitbrent/PptxGenJS/issues/42) ([priyaraskar](https://github.com/priyaraskar))
- Enhance bullet feature: offer diff types of bullets and add numbering option [\#49](https://github.com/gitbrent/PptxGenJS/issues/49) ([gitbrent](https://github.com/gitbrent))
- Add 4 new Presentation properties: `author`, `company`, `revision`, `subject` [\#53](https://github.com/gitbrent/PptxGenJS/pull/53) ([ericwgreene](https://github.com/ericwgreene))
- Moved to semver (semantic versioning)

## [v1.2.1](https://github.com/gitbrent/pptxgenjs/tree/v1.2.1) (2017-02-26)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.2.0...v1.2.1)

**Fixed Bugs:**
- Fixed issue with using percentages with `x`,`y`,`w`,`h` in `addTable()`
- Table formatting bug with rowspans and colspans [\#46](https://github.com/gitbrent/PptxGenJS/issues/46) ([itskun](https://github.com/itskun))

**Implemented Enhancements:**
- Allow more than a single 'x' and/or 'y' table location during Table Paging [\#43](https://github.com/gitbrent/PptxGenJS/issues/43) ([jenkinsns](https://github.com/jenkinsns))
- Bullets do not work with text objects in addText() method [\#44](https://github.com/gitbrent/PptxGenJS/issues/44) ([ellisgl](https://github.com/ellisgl))
- Table location and pagination [\#47](https://github.com/gitbrent/PptxGenJS/issues/47) ([itskun](https://github.com/itskun))
- Meta: Improve auto-paging in 'addTable()' [\#48](https://github.com/gitbrent/PptxGenJS/issues/48) ([gitbrent](https://github.com/gitbrent))
- Created a new common file (`pptxgenjs-demo.js`) to hold all demo code - now used by both the browser and the node demos.

## [v1.2.0](https://github.com/gitbrent/pptxgenjs/tree/v1.2.0) (2017-02-15)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.1.6...v1.2.0)

**Implemented Enhancements:**
- Pagination for `slideObj.addTable()`? [\#21](https://github.com/gitbrent/PptxGenJS/issues/21) ([TheDorkSide74](https://github.com/TheDorkSide74))
- Add support for media (Audio,Video,YouTube) [\#26](https://github.com/gitbrent/PptxGenJS/issues/26) ([shashank2104](https://github.com/shashank2104))
- How to set text shadow? [\#28](https://github.com/gitbrent/PptxGenJS/issues/28) ([itskun](https://github.com/itskun))
- Allow custom Layout sizes (ex: A3) [\#29](https://github.com/gitbrent/PptxGenJS/issues/29) ([itskun](https://github.com/itskun))
- Table cell marginPt should allow zero and take TRBL array [\#32](https://github.com/gitbrent/PptxGenJS/issues/32) ([ellisgl](https://github.com/ellisgl))
- Formatting rules do not apply to string with '\n' in `addText()` [\#34](https://github.com/gitbrent/PptxGenJS/issues/34) ([itskun](https://github.com/itskun))
- Node module appends to last generated PPT on `save()` [\#38](https://github.com/gitbrent/PptxGenJS/issues/38) ([alexanderpepper](https://github.com/alexanderpepper))
- callback support for save method [\#40](https://github.com/gitbrent/PptxGenJS/issues/40) ([ellisgl](https://github.com/ellisgl))
- Callback for save method (nodejs only) [\#41](https://github.com/gitbrent/PptxGenJS/pull/41) ([ellisgl](https://github.com/ellisgl))

**Fixed Bugs:**
- Table formatting bug in `addTable()` [\#36](https://github.com/gitbrent/PptxGenJS/issues/36) ([itskun](https://github.com/itskun))

## [v1.1.6](https://github.com/gitbrent/pptxgenjs/tree/v1.1.6) (2017-01-19)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.1.5...v1.1.6)

**Implemented Enhancements:**
- Support for animated GIFs in `addImage()` [\#22](https://github.com/gitbrent/PptxGenJS/issues/22) ([shashank2104](https://github.com/shashank2104))
- Added new `slideNumber` option allowing `x` and `y` placement of slide number [\#25](https://github.com/gitbrent/PptxGenJS/issues/25) ([priyaraskar](https://github.com/priyaraskar))

## [v1.1.5](https://github.com/gitbrent/pptxgenjs/tree/v1.1.5) (2017-01-17)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.1.4...v1.1.5)

**Fixed Bugs:**
- Trouble running in NW.js [\#19](https://github.com/gitbrent/PptxGenJS/issues/19) ([GregReser](https://github.com/GregReser))
- Supported usage via node program instead of HTML [\#23](https://github.com/gitbrent/PptxGenJS/issues/23) ([parsleyt](https://github.com/parsleyt))

## [v1.1.4](https://github.com/gitbrent/pptxgenjs/tree/v1.1.4) (2017-01-04)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.1.3...v1.1.4)

**Fixed Bugs:**
- Table formatting options set to default on empty cells [\#20](https://github.com/gitbrent/PptxGenJS/issues/20) ([rikvdk](https://github.com/rikvdk))
- Fixed issue with `addTable()` where passing "#" before hex value for `color` or `fill` option would generate an invalid slide

## [v1.1.3](https://github.com/gitbrent/pptxgenjs/tree/v1.1.3) (2016-12-28)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.1.2...v1.1.3)

**Implemented Enhancements:**
- Add new options to `addSlidesForTable()` allowing for placement and size: `x`,`y`,`w`,`h` [\#18](https://github.com/gitbrent/PptxGenJS/issues/18) ([priyaraskar](https://github.com/priyaraskar))

**Fixed Bugs:**
- Cannot read property 'opts' of null [\#17](https://github.com/gitbrent/PptxGenJS/issues/17) ([ninas880025](https://github.com/ninas880025))

## [v1.1.2](https://github.com/gitbrent/pptxgenjs/tree/v1.1.2) (2016-12-16)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.1.1...v1.1.2)

**Implemented Enhancements:**
- The Slide `addTable()` method was modified to reduce the options passed from 2 objects to a single one

**Fixed Bugs:**
- The colW `addTable()` option is not working [\#15](https://github.com/gitbrent/PptxGenJS/issues/15) ([ninas880025](https://github.com/ninas880025))
- Modified `addSlidesForTable()`: table selectors made more specific by selecting only direct children now (nested tables would cause excessive looping) [\#14](https://github.com/gitbrent/PptxGenJS/issues/14) ([forrahul123](https://github.com/forrahul123))
- Fixed crash caused by calling `addText` without an options object

## [v1.1.1](https://github.com/gitbrent/pptxgenjs/tree/v1.1.1) (2016-12-08)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.1.0...v1.1.1)

**Implemented Enhancements:**
- Major documentation update
- Added instructions to `pptxgenjs.masters.js` file, plus more examples and code
- Added sandbox/ad-hoc code area to demo page

**Fixed Bugs:**
- Table with 7 columns generates an invalid pptx file [\#12](https://github.com/gitbrent/PptxGenJS/issues/12) ([rikvdk](https://github.com/rikvdk))

## [v1.1.0](https://github.com/gitbrent/pptxgenjs/tree/v1.1.0) (2016-11-22)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.0.1...v1.1.0)

**Implemented Enhancements:**
- Added support for base64-encoded images
- Adding npm dependencies [\#4](https://github.com/gitbrent/PptxGenJS/pull/1) ([DzmitryDulko](https://github.com/DzmitryDulko))
- Added support for italic text [\#6](https://github.com/gitbrent/PptxGenJS/issues/6) ([stevenljacobsen](https://github.com/stevenljacobsen))
- Added ability to selectively override Master Slide background color/image [\#7](https://github.com/gitbrent/PptxGenJS/issues/7) ([stevenljacobsen](https://github.com/stevenljacobsen))
- How can customize pptx theme? [\#9](https://github.com/gitbrent/PptxGenJS/issues/9) ([ielijose](https://github.com/ielijose))
- Add Rectangle to supported Master Slide shapes [\#10](https://github.com/gitbrent/PptxGenJS/pull/10) ([ielijose](https://github.com/ielijose))
- Added support for bulleted text [\#11](https://github.com/gitbrent/PptxGenJS/issues/11) ([gojko](https://github.com/gojko))

**Fixed Bugs:**
- Fix repo URL in package.json [\#5](https://github.com/gitbrent/PptxGenJS/pull/5) ([pdehaan](https://github.com/pdehaan))

## [v1.0.1](https://github.com/gitbrent/pptxgenjs/tree/v1.0.1) (2016-09-03)
[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.0.0...v1.0.1)

**Implemented enhancements:**
- Moved from `cx` and `cy` option keys to `w` and `h`
- Adding ability to load data uri as images/Updating jszip library [\#2](https://github.com/gitbrent/PptxGenJS/pull/2) ([DzmitryDulko](https://github.com/DzmitryDulko))
- Publish library as npm package [\#3](https://github.com/gitbrent/PptxGenJS/issues/3) ([DzmitryDulko](https://github.com/DzmitryDulko))

**Fixed Bugs:**
- Fixed resource references [\#1](https://github.com/gitbrent/PptxGenJS/pull/1) ([DzmitryDulko](https://github.com/DzmitryDulko))

## [v1.0.0](https://github.com/gitbrent/pptxgenjs/tree/v1.0.0) (2016-03-29)

**Initial Release**
