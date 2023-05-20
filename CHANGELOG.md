# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Notes

## [3.13.0](https://github.com/gitbrent/PptxGenJS/releases/tag/v3.13.0) - 2023-0?-0?

- Added `textDirection` property for text and table cells to allow vertical rotation of text ([gitbrent](https://github.com/gitbrent))

### Changed

- Bump jszip to ^3.10.1 [\#1255](https://github.com/gitbrent/PptxGenJS/pull/1255) ([NateRadebaugh](https://github.com/NateRadebaugh))

## [3.12.0](https://github.com/gitbrent/PptxGenJS/releases/tag/v3.12.0) - 2023-03-19

### Added

- Added selecting round or square line cap on line charts [\#1126](https://github.com/gitbrent/PptxGenJS/pull/1126) ([mathbruyen](https://github.com/mathbruyen))
- Added `newAutoPagedSlides` method to `slide` (resolves issue #625) [\#1133](https://github.com/gitbrent/PptxGenJS/pull/1133) ([mikemeerschaert](https://github.com/mikemeerschaert))
- Added optional image shadow props [\#1147](https://github.com/gitbrent/PptxGenJS/pull/1147) ([seekuehe](https://github.com/seekuehe))
- Added ability to set default fontFace [\#1158](https://github.com/gitbrent/PptxGenJS/issues/1158) ([matt88120](https://github.com/matt88120))

### Fixed

- Fixed `autoPage` duplicates text when text array is used [\#1139](https://github.com/gitbrent/PptxGenJS/issues/1139) ([mikemeerschaert](https://github.com/mikemeerschaert))
- PowerPoint shows the "repair" dialog when adding an SVG image to a slide master [\#1150](https://github.com/gitbrent/PptxGenJS/issues/1150) ([BenHall-1](https://github.com/BenHall-1))
- Fixed gh-pages text api docs: transparency + wrap [\#1153](https://github.com/gitbrent/PptxGenJS/pull/1153) ([tjinauyeung](https://github.com/tjinauyeung))
- Fixed YouTube videos not working [\#1156](https://github.com/gitbrent/PptxGenJS/issues/1156) ([gitbrent](https://github.com/gitbrent))
- Fixed handle `holeSize=0` for doughnut chart [\#1180](https://github.com/gitbrent/PptxGenJS/pull/1180) ([mathbruyen](https://github.com/mathbruyen))
- Fixed 3D chart options not working correctly (and updated demo) ([gitbrent](https://github.com/gitbrent))

### Changed

- (Internal) migrate library from tslint to eslint [\#1155](https://github.com/gitbrent/PptxGenJS/pull/1155) ([gitbrent](https://github.com/gitbrent))

## [3.11.0] - 2022-08-06

### Added

- Added category crosses at property (`catAxisCrossesAt`) [\#966](https://github.com/gitbrent/PptxGenJS/pull/966) ([parvezapathan](https://github.com/parvezapathan))
- Added support for multi-level category axes [\#1012](https://github.com/gitbrent/PptxGenJS/pull/1012) ([MariusOpeepl](https://github.com/MariusOpeepl))
- Added 2 new Chart props: `plotArea` and `chartArea` allowing fill and border for each (`plotArea` deprecates `fill` and `border`) [\#1015](https://github.com/gitbrent/PptxGenJS/issues/1015) ([hvstaden](https://github.com/hvstaden))
- Added serie name on bubble chart, category axis position, leader lines on bubble chart [\#1100](https://github.com/gitbrent/PptxGenJS/pull/1100) ([mathbruyen](https://github.com/mathbruyen))
- Added `bubble3D` chart type [\#1108](https://github.com/gitbrent/PptxGenJS/pull/1108) ([mathbruyen](https://github.com/mathbruyen))
- Added new tool under demos: `data_convert` which turns Excel (tab-delim) data to chart data type easily ([gitbrent](https://github.com/gitbrent))

### Fixed

- Using `addImage()` with uppercase path prop causes "needs to repair presentation" [\#860](https://github.com/gitbrent/PptxGenJS/issues/860) ([mamodo123](https://github.com/mamodo123))
- Chart with lines and bars produces repair file dialog in Powerpoint [\#1013](https://github.com/gitbrent/PptxGenJS/issues/1013) ([kornarakis](https://github.com/kornarakis))
- Bubble Charts limited to 26 columns [\#1076](https://github.com/gitbrent/PptxGenJS/issues/1076) ([benjaminpavone](https://github.com/benjaminpavone))
- Using `addImage` with `tableToSlides()` does not work [\#1103](https://github.com/gitbrent/PptxGenJS/issues/1103) ([Strawberry0215](https://github.com/Strawberry0215))
- escape object name in chart xml [\#1122](https://github.com/gitbrent/PptxGenJS/pull/1122) ([mathbruyen](https://github.com/mathbruyen))
- Several issues with charts embedded Excel sheets that prevented "Edit Data in Excel" from working ([gitbrent](https://github.com/gitbrent))
- Issue with combo charts secondary axis on wrong side ([gitbrent](https://github.com/gitbrent))
- Issue with chart prop `titlePos` not working ([gitbrent](https://github.com/gitbrent))

### Changed

- react-demo: updated `react-scripts` to v5.0.0 from v4 ([gitbrent](https://github.com/gitbrent))

## [3.10.0] - 2022-04-10

### Added

- Add name (`objectName`) to all core objects [\#1019](https://github.com/gitbrent/PptxGenJS/pull/1019) ([mvecsernyes](https://github.com/mvecsernyes))
- Add image transparency [\#1053](https://github.com/gitbrent/PptxGenJS/pull/1053) ([mmarkelov](https://github.com/mmarkelov))
- Add text transparency [\#1054](https://github.com/gitbrent/PptxGenJS/issues/1054) ([ibrahimovfuad](https://github.com/ibrahimovfuad))

### Fixed

- Radar chart line colors [\#539](https://github.com/gitbrent/PptxGenJS/issues/539) ([pablodicosta](https://github.com/pablodicosta))
- Placeholder definitions missing props [\#987](https://github.com/gitbrent/PptxGenJS/issues/987) ([bigbug](https://github.com/bigbug))
- Charts and media together is causing pptx needs repair error [\#1020](https://github.com/gitbrent/PptxGenJS/issues/1020) ([mvecsernyes](https://github.com/mvecsernyes))
- Adding hyperlink to table cell doesn't work [\#1049](https://github.com/gitbrent/PptxGenJS/issues/1049) ([tbowmo](https://github.com/tbowmo))
- Underline doesn't work in table after update to v3.9.0 [\#1052](https://github.com/gitbrent/PptxGenJS/issues/1052) ([hhq365](https://github.com/hhq365))
- `ImageProps.sizing` props `w`, `h`, `x`, `y` s/b typed `Coord` [\#1065](https://github.com/gitbrent/PptxGenJS/issues/1065) ([Naveencheekoti17](https://github.com/BistroStu))
- `ImageProps.sizing` are type Coord [\#1066](https://github.com/gitbrent/PptxGenJS/pull/1066) ([BistroStu](https://github.com/BistroStu))
- `transparency` doesn't work in table cell [\#1095](https://github.com/gitbrent/PptxGenJS/issues/1095) ([pipipi-pikachu](https://github.com/pipipi-pikachu))

## [3.9.0] - 2021-12-11

### Added

- Added overlap parameter to bar charts [\#1010](https://github.com/gitbrent/PptxGenJS/pull/1010) ([Norfaer](https://github.com/Norfaer))
- Slide number can now be set as bold [\#1016](https://github.com/gitbrent/PptxGenJS/pull/1016) ([mathbruyen](https://github.com/mathbruyen))
- Added media cover images & file extensions; media is reused now (same file only loaded/written once) [\#1024](https://github.com/gitbrent/PptxGenJS/pull/1024) ([canwdev](https://github.com/canwdev))

### Fixed

- Use `encodeXmlEntities()` for formatCode attributes [\#955](https://github.com/gitbrent/PptxGenJS/pull/955) ([dimfeld](https://github.com/dimfeld))
- SlideNumber vertical alignment (`valign`) not working [\#1000](https://github.com/gitbrent/PptxGenJS/pull/1000) ([kramsram](https://github.com/kramsram))
- Fix for InvertedColors (Issue #970) [\#1004](https://github.com/gitbrent/PptxGenJS/pull/1004) ([leonyah](https://github.com/leonyah))
- PPT repair issue for long text [\#1008](https://github.com/gitbrent/PptxGenJS/issues/1008) ([Naveencheekoti17](https://github.com/Naveencheekoti17)), fixed via [\#1028](https://github.com/gitbrent/PptxGenJS/pull/1028) ([gitbrent](https://github.com/gitbrent))
- Doughnut chart: each data marker as a different color [\#1017](https://github.com/gitbrent/PptxGenJS/pull/1017) ([mathbruyen](https://github.com/mathbruyen))

### Changed

- React Demo: updated to latest create-react-app ([gitbrent](https://github.com/gitbrent))

## [3.8.0] - 2021-09-28

### Added

- Table auto-paging completely re-written from scratch; finally handles complex-text (text runs) [\#993](https://github.com/gitbrent/PptxGenJS/pull/993) ([gitbrent](https://github.com/gitbrent))

### Changed

- Browser Demo: refreshed UI and upgraded to bootstrap-5 [\#997](https://github.com/gitbrent/PptxGenJS/pull/997) ([gitbrent](https://github.com/gitbrent))
- Documentation site (gh-pages) rebuilt from scratch [\#999](https://github.com/gitbrent/PptxGenJS/pull/999) ([gitbrent](https://github.com/gitbrent))

## [3.7.1] - 2021-07-21

### Fixed

- Added missing `altText` prop to ImageProps [\#848](https://github.com/gitbrent/PptxGenJS/pull/848) ([yorch](https://github.com/yorch))

## [3.7.0] - 2021-07-20

### Added

- Alt Text to images [\#848](https://github.com/gitbrent/PptxGenJS/pull/848) ([yorch](https://github.com/yorch))
- Custom geometry support (freeform) [\#872](https://github.com/gitbrent/PptxGenJS/pull/872) ([apresmoi](https://github.com/apresmoi))
  - Resolves:
    - Custom polygon generation [\#597](https://github.com/gitbrent/PptxGenJS/issues/597) ([hirenj](https://github.com/hirenj))
    - Is there any way to draw a bell curve shape? [\#946](https://github.com/gitbrent/PptxGenJS/issues/946) ([gurdeep-sourcefuse](https://github.com/gurdeep-sourcefuse))

### Fixed

- Background in master template broken (support multiple `background` props) [\#968](https://github.com/gitbrent/PptxGenJS/issues/968) ([viral-sh](https://github.com/viral-sh))
- Arguments for radius not allowed in TypeScript for rectangles [\#969](https://github.com/gitbrent/PptxGenJS/issues/969) ([ln56b](https://github.com/ln56b))
- Documentation: `catAxisLine*` and `valAxisLine*` props missing [\#980](https://github.com/gitbrent/PptxGenJS/issues/980) ([ln56b](https://github.com/hhq365))

### Chart Updates

Comprehensive Pull

- Multiple Chart Enhancements and Bugfixes [\#938](https://github.com/gitbrent/PptxGenJS/pull/938) ([ReimaFrgos](https://github.com/ReimaFrgos))
  - Resolves:
    - Using scheme colors and fonts in chart axis labels, axis lines and series labels #858 [robertedjones]
    - dataLabelPosition option for Pie charts #837 [kornarakis]
    - Bubble chart catAxisMajorUnit not working #747 [dscdngnw]
    - dataLabelFontBold option not working as expected. #662 [belall-shaikh]
    - dataLabelPosition is not working in Multi Type Charts #815 [Adt-SakshamSethi]
    - dataLabelPosition "t" in Bar chart is crashing ppt in latest MS office Power Point #788 [jsvishal]
    - Setting dataLabelPosition to a line chart causes latest office application to ask for repair #768 [artdomg]

## [3.6.0] - 2021-05-02

### Release Summary

- **Major Update**: demo code (they're all .mjs modules now!); dropped support for IE11 (RIP!) in demo app.
- **IE11 Note**: Dropped support for IE11 (use v3.5.0 or below) (library still works with IE11 using polyfill)

### Added

- Alt Text to charts [\#848](https://github.com/gitbrent/PptxGenJS/pull/848) ([yorch](https://github.com/yorch))
- Tab Stops to Text objects [\#853](https://github.com/gitbrent/PptxGenJS/pull/853) ([wangfengming](https://github.com/wangfengming))
- Text Highlight to Text objects [\#857](https://github.com/gitbrent/PptxGenJS/pull/857) ([wangfengming](https://github.com/wangfengming))
- Transparency to line [\#889](https://github.com/gitbrent/PptxGenJS/pull/889) ([mmarkelov](https://github.com/mmarkelov))
- Transparency to slide [\#891](https://github.com/gitbrent/PptxGenJS/pull/891) ([mmarkelov](https://github.com/mmarkelov))

### Changed

- Website/Docs Docusaurus v2.0; major UI facelift [\#931](https://github.com/gitbrent/PptxGenJS/pull/931) ([gitbrent](https://github.com/gitbrent))

### Deprecated

- Slide.fill (`BackgroundProps`) - use `ShapeFillProps` instead

### Removed

- Browser Demo: Dropped support for IE11 (use v3.5.0 or below) (library still works with IE11 using polyfill)

### Fixed

- Margin not working with placeholder text [\#640](https://github.com/gitbrent/PptxGenJS/issues/640) ([bestis](https://github.com/bestis))
- Cant create a list of bulleted links in a table cell [\#763](https://github.com/gitbrent/PptxGenJS/issues/763) ([avillamaina](https://github.com/avillamaina))
- Small API documentation glitch [\#895](https://github.com/gitbrent/PptxGenJS/issues/895) ([Slidemagic](https://github.com/Slidemagic))
- pptx.stream() WriteBaseProps should be optional [\#932](https://github.com/gitbrent/PptxGenJS/issues/932) ([arbourd](https://github.com/arbourd))
- Running StdTests generate a corrupt PPT [\#937](https://github.com/gitbrent/PptxGenJS/issues/937) ([michaeltford](https://github.com/michaeltford))
- addNotes function adding notes as an array of objects, parsed as [object Object] in notes field [\#941](https://github.com/gitbrent/PptxGenJS/issues/941) ([karlolsonuc](https://github.com/karlolsonuc))

## [3.5.0] - 2021-03-30

### Release Summary

- write()/writeFile() method string arguments are deprecated - props object in now the sole arg (`WriteProps`/`WriteFileProps`)

### Added

- Enabled JSZip compression [\#713](https://github.com/gitbrent/PptxGenJS/issues/713) ([pimlottc-gov](https://github.com/pimlottc-gov))
- Soft line break property: `softBreakBefore` [\#806](https://github.com/gitbrent/PptxGenJS/pull/806) ([memorsolutions](https://github.com/memorsolutions))
- More text styles: underline/strike/baseline [\#854](https://github.com/gitbrent/PptxGenJS/pull/854) ([wangfengming](https://github.com/wangfengming))
- Support line spacing by multiple: `lineSpacingMultiple` [\#855](https://github.com/gitbrent/PptxGenJS/pull/855) ([wangfengming](https://github.com/wangfengming))
- Chart val axis option: logarithmic scale base: `valAxisLogScaleBase` [\#878](https://github.com/gitbrent/PptxGenJS/issues/878) ([rkspx](https://github.com/rkspx))

### Changed

- Fixed: Setting the "Wrap text in shape" option [\#771](https://github.com/gitbrent/PptxGenJS/issues/771) ([CroniD](https://github.com/CroniD))
- Fixed: `dataLabelFormatCode` option creates corrupted file if the value includes quotes [\#834](https://github.com/gitbrent/PptxGenJS/issues/834) ([kornarakis](https://github.com/kornarakis)) [\#884](https://github.com/gitbrent/PptxGenJS/pull/884) ([gazlo](https://github.com/gazlo))
- Fixed: Improve typescipt defs: fix dupes, etc [\#886](https://github.com/gitbrent/PptxGenJS/pull/886) ([mmarkelov](https://github.com/mmarkelov))
- Fixed: Wrong type definition for placeholder type property [\#921](https://github.com/gitbrent/PptxGenJS/issues/921) ([lukevella](https://github.com/lukevella))

### Internal Updates

- Doc/Website Updates: Docusaurus docs and website updated to v2.0 [\#924](https://github.com/gitbrent/PptxGenJS/pull/924) ([gitbrent](https://github.com/gitbrent))

## [3.4.0] - 2021-01-03

### Added

- Added: `firstSliceAngle` (Pie, Doughnut charts) [\#666](https://github.com/gitbrent/PptxGenJS/issues/666) ([ghost](https://github.com/ghost)) [\#809](https://github.com/gitbrent/PptxGenJS/pull/809) ([cronin4392](https://github.com/cronin4392))
- Added: Ability to change hyperlink `color` [\#389](https://github.com/gitbrent/PptxGenJS/issues/389) ([szilagyikinga](https://github.com/szilagyikinga)) [\#793](https://github.com/gitbrent/PptxGenJS/pull/793) ([ReimaFrgos](https://github.com/ReimaFrgos))
- Added: Horizontal/Vertical flip capability to images [\#824](https://github.com/gitbrent/PptxGenJS/pull/824) ([luism-s](https://github.com/luism-s))
- Added: New `titleBold` option on chart settings [\#830](https://github.com/gitbrent/PptxGenJS/pull/830) ([twatson83](https://github.com/twatson83))
- Added: New cat/val-AxisLineColor/AxisLineSize/AxisLineStyle chart options [\#831](https://github.com/gitbrent/PptxGenJS/pull/831) ([twatson83](https://github.com/twatson83))
- Added: New shape options: `angleRange` and `arcThicknessRatio` [\#547](https://github.com/gitbrent/PptxGenJS/issues/547) ([paolochiodi](https://github.com/paolochiodi)) [\#861](https://github.com/gitbrent/PptxGenJS/pull/861) ([apresmoi](https://github.com/apresmoi))

### Changed

- Fixed: catAxisLabelPos and valAxisLabelPos options are not working [\#709](https://github.com/gitbrent/PptxGenJS/issues/709) ([cpf121](https://github.com/cpf121))
- Fixed: logic for dataLabelFormat code in Pie and Donut charts [\#802](https://github.com/gitbrent/PptxGenJS/pull/802) ([cronin4392](https://github.com/cronin4392))
- Fixed: data label position for Pie chart [\#808](https://github.com/gitbrent/PptxGenJS/pull/808) ([cronin4392](https://github.com/cronin4392))
- Fixed: Single data set with a custom color should not create legends for each category [\#821](https://github.com/gitbrent/PptxGenJS/issues/821) ([tvt](https://github.com/tvt))
- Fixed: bug when evaluating `catAxisLabelPos`,`valAxisLabelPos` props [\#829](https://github.com/gitbrent/PptxGenJS/pull/829) ([twatson83](https://github.com/twatson83))
- Fixed: secondary axis param (`secondaryValAxis`) check [\#832](https://github.com/gitbrent/PptxGenJS/pull/832) ([twatson83](https://github.com/twatson83))
- Fixed: `addSection` method missing return type in `index.d.ts` [\#833](https://github.com/gitbrent/PptxGenJS/issues/833) ([dylang](https://github.com/dylang))
- Fixed: Align property doesn't work in slide number object [\#835](https://github.com/gitbrent/PptxGenJS/issues/835) ([ax2mx](https://github.com/ax2mx))
- Fixed: Margin doesn't work in slide number object [\#836](https://github.com/gitbrent/PptxGenJS/issues/836) ([ax2mx](https://github.com/ax2mx))
- Fixed: several rounding mistakes for precision, accuracy, and usability [\#840](https://github.com/gitbrent/PptxGenJS/pull/840) ([michaelcbrook](https://github.com/michaelcbrook))
- Fixed: catAxisMinorTickMark [\#841](https://github.com/gitbrent/PptxGenJS/pull/841) ([twatson83](https://github.com/twatson83))
- Fixed: colspan/rowspan [\#852](https://github.com/gitbrent/PptxGenJS/pull/852) ([wangfengming](https://github.com/wangfengming))
- Fixed: typo in ts doc [\#873](https://github.com/gitbrent/PptxGenJS/issues/873) ([jencii](https://github.com/jencii))
- Fixed: TypeError: Cannot set property 'lIns' of undefined [\#879](https://github.com/gitbrent/PptxGenJS/issues/879) ([CroniD](https://github.com/CroniD))

### Internal Updates

- Library Updates: TypeScript 4, Rollup 2.3 and more [\#866](https://github.com/gitbrent/PptxGenJS/pull/866) ([gitbrent](https://github.com/gitbrent))

## [3.3.1] - 2020-08-23

### Changed

- Fixed: Broken pptx has generated if used custom slide layout in v3.3.0 [\#826](https://github.com/gitbrent/PptxGenJS/issues/826) ([yhatt](https://github.com/yhatt))
- Fixed: lineSpacing option set to decimal triggers repair alert [\#827](https://github.com/gitbrent/PptxGenJS/issues/827) ([ReimaFrgos](https://github.com/ReimaFrgos))
- Updated `demos.js` to replace all fill:string with fill:ShapeFillProps ([gitbrent](https://github.com/gitbrent))

## [3.3.0] - 2020-08-16

### Major Change Summary

- The `addTable()` method finally supports auto-paging, including support for repeating table headers!
- The `addText()` method text layout engine has been rewritten from scratch and handles every type of layout case now
- New `addText()` `fit` option ('none' | 'shrink' | 'resize') addresses long-standing issues with shrink/resize objects (new demo page as well)
- Fix for Angular "`Buffer` is unknown" issue
- Major update of typescript defs, including tons of documentation that has been added
- Unfotunately, `fill` no longer accepts a plain string and there was no smooth way to make that backwards compatible (sorry!)

### BREAKING CHANGES

- **TypeScript users**: `fill` property no longer accepts strings, only `ShapeFill` type now (sorry!)
- **All users**: table and textbox text linebreaks may act differently! (a major rewrite to correct long-standing issues with alignment/breakLine finally landed)

### Added

- Added: Auto-Paging finally comes to `addTable()` [\#262](https://github.com/gitbrent/PptxGenJS/issues/262) ([okaiyong](https://github.com/okaiyong))
- Added: Chart DataTable formatting `dataTableFormatCode` and `valLabelFormatCode` [\#489](https://github.com/gitbrent/PptxGenJS/issues/489) ([phobos7000](https://github.com/phobos7000)) [\#684](https://github.com/gitbrent/PptxGenJS/pull/684) ([hanzi](https://github.com/hanzi))
- Added: Background image for slides (deprecated `bkgd:string` with `background:BkgdOpts`) [\#610](https://github.com/gitbrent/PptxGenJS/pull/610) ([thomasowow](https://github.com/thomasowow))
- Added: `shapeName` to objects instead of default [\#724](https://github.com/gitbrent/PptxGenJS/issues/724) ([Offbeatmammal](https://github.com/Offbeatmammal))
- Added: `valAxisDisplayUnitLabel` option [\#765](https://github.com/gitbrent/PptxGenJS/pull/765) ([hysh](https://github.com/hysh))
- Added: Ability to create a hyperlink on a shape [\#767](https://github.com/gitbrent/PptxGenJS/issues/767) ([CroniD](https://github.com/CroniD))

### Changed

- Fixed: complete rewrite of genXmlTextBody for new text run/paragraph generation. Fixes: [\#369](https://github.com/gitbrent/PptxGenJS/issues/369)
  [\#448](https://github.com/gitbrent/PptxGenJS/issues/448), [\#460](https://github.com/gitbrent/PptxGenJS/issues/460), [\#751](https://github.com/gitbrent/PptxGenJS/issues/751), [\#772](https://github.com/gitbrent/PptxGenJS/pull/772)
- Fixed: tableToSlides `addHeaderToEach` finally duplicates all header rows, not just the first one [\#262](https://github.com/gitbrent/PptxGenJS/issues/262) ([okaiyong](https://github.com/okaiyong))
- Fixed `colW` length mismatch with colspans (Issue #651) [\#679](https://github.com/gitbrent/PptxGenJS/issues/679) ([Joshua-rose](https://github.com/Joshua-rose))
- Fixed: hyperlink and tooltip property `rId` is not working? [\#758](https://github.com/gitbrent/PptxGenJS/issues/758) ([kuldeept70](https://github.com/kuldeept70))
- Fixed: removed old/unused options from demo [\#759](https://github.com/gitbrent/PptxGenJS/pull/759) ([sijmenvos](https://github.com/sijmenvos))
- Fixed: removed `Buffer` type from `index.ts.d` [\#761](https://github.com/gitbrent/PptxGenJS/pull/761) ([lustigerlurch551](https://github.com/lustigerlurch551))
- Fixed: addSection does not escape XML unsafe characters [\#774](https://github.com/gitbrent/PptxGenJS/issues/774) ([pimlottc-gov](https://github.com/pimlottc-gov))
- Fixed: Multiple Border Types not supported in Table Cell [\#775](https://github.com/gitbrent/PptxGenJS/issues/775) ([jsvishal](https://github.com/jsvishal))
- Fixed: New ITextOpts `fit` prop, removed `autoFit`/`shrinkText`, new demo slide [\#779](https://github.com/gitbrent/PptxGenJS/issues/779) ([DonnaZukowskiPfizer](https://github.com/DonnaZukowskiPfizer)) ([ReimaFrgos](https://github.com/ReimaFrgos))
- Fixed: EMU calculations are not safe (calcPointValue in gen-xml) [\#781](https://github.com/gitbrent/PptxGenJS/issues/781) ([CroniD](https://github.com/CroniD))
- Fixed: type defs for `TableCell.text` not correct ([gitbrent](https://github.com/gitbrent))
- Fixed: type defs for `ITableOptions` s/b `TableOptions` ([gitbrent](https://github.com/gitbrent))

## [3.2.1] - 2020-05-25

### Added

### Changed

- Fixed: `addTable`, `addText`, etc. not working properly inside tableToSlides [\#715](https://github.com/gitbrent/PptxGenJS/issues/715) ([Smithvinayakiya](https://github.com/Smithvinayakiya))
- Fixed: Issue links in release notes are broken [\#749](https://github.com/gitbrent/PptxGenJS/issues/749) ([pimlottc-gov](https://github.com/pimlottc-gov))
- Fixed: Type defs were missing ISlideMasterOptions `text` prop and `slideNumber` align ([gitbrent](https://github.com/gitbrent))
- Fixed: Type defs misspelled `rowW` s/b `rowH` ([gitbrent](https://github.com/gitbrent))
- Fixed: Documentation: Corrected max value for `barGapWidthPct` ([gitbrent](https://github.com/gitbrent))

## [3.2.0] - 2020-05-17

### Added

- Added: New chart type: Stacked Area Charts [\#333](https://github.com/gitbrent/PptxGenJS/issues/333) ([fordaaronj](https://github.com/fordaaronj))
- Added: Sections can now be created [\#349](https://github.com/gitbrent/PptxGenJS/issues/349) ([atulsingh0913](https://github.com/atulsingh0913))
- Added: New bullet option `marginPt` to control left indent margin [\#504](https://github.com/gitbrent/PptxGenJS/issues/504) ([Cyan005](https://github.com/Cyan005))

### Changed

- Fixed: Placeholder type Body is defaulting in a hanging indent [\#589](https://github.com/gitbrent/PptxGenJS/issues/589) ([colmben](https://github.com/colmben))
- Fixed: Text in slides does not override the bullet master [\#620](https://github.com/gitbrent/PptxGenJS/pull/620) ([sgenoud](https://github.com/sgenoud))
- Fixed: Type errors in `index.d.ts` [\#672](https://github.com/gitbrent/PptxGenJS/issues/672) ([Krishnakanth94](https://github.com/Krishnakanth94))
- Fixed: Typescript defs Slide and ISlide [\#673](https://github.com/gitbrent/PptxGenJS/issues/673) ([gytisgreitai](https://github.com/gytisgreitai))
- Fixed: Spelling consistent "Presenation" -> "Presentation" typo [\#694](https://github.com/gitbrent/PptxGenJS/pull/694) ([ankon](https://github.com/ankon))
- Fixed: Handle errors with promise rejections [\#695](https://github.com/gitbrent/PptxGenJS/pull/695) ([ankon](https://github.com/ankon))
- Fixed: Update 'pptx' to 'pres' in README.md [\#700](https://github.com/gitbrent/PptxGenJS/pull/700) ([lucidlemon](https://github.com/lucidlemon))
- Fixed: Time units validation [\#706](https://github.com/gitbrent/PptxGenJS/pull/706) ([lucasflomuller](https://github.com/lucasflomuller))
- Fixed: Add the slide layout name to the generated background image name [\#726](https://github.com/gitbrent/PptxGenJS/pull/726) ([jrohland](https://github.com/jrohland))
- Fixed: Type issue addTable rows, updated TableCell/TableRow [\#735](https://github.com/gitbrent/PptxGenJS/issues/735) ([robertsoaa](https://github.com/robertsoaa))
- Continued improvement of typescript definitions file ([gitbrent](https://github.com/gitbrent))

## [3.1.1] - 2020-02-02

### Added

- TypeScript: Add shapes and font options types [\#650](https://github.com/gitbrent/PptxGenJS/pull/650) ([cronin4392](https://github.com/cronin4392))
- TypeScript: Added correct export of types and ts-def file (`pptx.ShapeType.rect`, etc) in `index.d.ts` ([gitbrent](https://github.com/gitbrent))

### Changed

- Fixed: Re-added "browser" property to `package.json` to avoid old "fs not found" Angular/webpack issue (Angular 8) [\#654](https://github.com/gitbrent/PptxGenJS/issues/654) ([cwilkens](https://github.com/cwilkens))
- Fixed: Previous release introduced a regression bug and broke addTest placeholder's ([gitbrent](https://github.com/gitbrent))
- Fixed: addChart and addImage in the same slide cause an error [fixed via `getNewRelId`] [\#655](https://github.com/gitbrent/PptxGenJS/issues/655) ([JuliaSheleva](https://github.com/JuliaSheleva))

### Removed

- The `core-shapes.ts` file was removed, shape def collapsed to simple type array, rolled into `core-enums.ts` and `index.d.ts` ([gitbrent](https://github.com/gitbrent))

## [3.1.0] - 2020-01-21

### Added

- Added `valAxisDisplayUnit` [\#606](https://github.com/gitbrent/PptxGenJS/pull/606) ([AmrutPatil](https://github.com/AmrutPatil))
- Added `dataTableFontSize` chart option [\#622](https://github.com/gitbrent/PptxGenJS/pull/622) ([MehdiAroui](https://github.com/MehdiAroui))
- Added text `glow` option [\#630](https://github.com/gitbrent/PptxGenJS/pull/630) ([kevinresol](https://github.com/kevinresol))
- Ability to `rotate` image [\#639](https://github.com/gitbrent/PptxGenJS/pull/639) ([alabaki](https://github.com/alabaki))
- Include types in package.json files [\#641](https://github.com/gitbrent/PptxGenJS/pull/641) ([cronin4392](https://github.com/cronin4392))
- Added `showLeaderLines` chart option [\#642](https://github.com/gitbrent/PptxGenJS/pull/642) ([cronin4392](https://github.com/cronin4392))

### Changed

- Fixed: Empty color negative values on barchart [\#285](https://github.com/gitbrent/PptxGenJS/issues/285) ([andrei-cs](https://github.com/andrei-cs)) ([Slidemagic](https://github.com/Slidemagic))
- Fixed: Add missing margin type from ITextOpts [\#643](https://github.com/gitbrent/PptxGenJS/pull/643) ([cronin4392](https://github.com/cronin4392))
- Fixed: Scatter plot `dataLabelPosition` [\#644](https://github.com/gitbrent/PptxGenJS/issues/644) ([afarghaly10](https://github.com/afarghaly10))
- Fixed: Added new babel polyfill for IE11; other IE11 fixes in demo, etc. [\#648](https://github.com/gitbrent/PptxGenJS/issues/648) ([YakQin](https://github.com/YakQin))
- Updated Demo: added support for light/dark mode; new Image slide for rotation; new busy progress modal ([gitbrent](https://github.com/gitbrent))

### Removed

- Removed: jsdom pkg is no longer a dependency in `package.json` ([gitbrent](https://github.com/gitbrent))

## [3.0.1] - 2020-01-07

### Changed

- Fixed: JSZip not found under Node.js [\#638](https://github.com/gitbrent/PptxGenJS/issues/638) ([rse](https://github.com/rse))
- Fixed: react demo fixes and new build for [demo-react online](https://gitbrent.github.io/PptxGenJS/demo-react/index.html) ([gitbrent](https://github.com/gitbrent))
- Fixed: added missing catch on media promise.all to handle 404 media links ([gitbrent](https://github.com/gitbrent))
- Fixed: replaced wikimedia links in common/demos.js with github raw content links ([gitbrent](https://github.com/gitbrent))

## [3.0.0] - 2020-01-01

### Added

- Ability to specify numbered list format [\#452](https://github.com/gitbrent/PptxGenJS/issues/452) ([mayvazyan](https://github.com/mayvazyan))
- New cat/val axis options: majorTickMark/minorTickMark [\#473](https://github.com/gitbrent/PptxGenJS/pull/473) ([RokasDie](https://github.com/RokasDie))
- Ability to set start number "startAt" for a bullet list of type numbered [\#554](https://github.com/gitbrent/PptxGenJS/issues/554) [\#555](https://github.com/gitbrent/PptxGenJS/pull/555) ([bj-mitchell](https://github.com/bj-mitchell))

### Changed

- Fixed: Set proper MIME type for PPTX presentation [\#471](https://github.com/gitbrent/PptxGenJS/issues/471) ([StefanBrand](https://github.com/StefanBrand))
- Fixed: SVG images used to be generated by Node [\#515](https://github.com/gitbrent/PptxGenJS/issues/515) ([michaelcbrook](https://github.com/michaelcbrook))
- Fixed: SVG support has several issues [\#528](https://github.com/gitbrent/PptxGenJS/pull/528) ([RicardoNiepel](https://github.com/RicardoNiepel))
- Fixed: Downloading PPT in iOS using Safari does not work. File named as UNKNOWN. [\#540](https://github.com/gitbrent/PptxGenJS/issues/540) ([mustafagentrit](https://github.com/mustafagentrit))
- Fixed: Tables not being displayed after update [\#559](https://github.com/gitbrent/PptxGenJS/issues/559) ([emartz404](https://github.com/emartz404))
- Fixed: Hyperlink creates malformed slide if it includes "&" [\#562](https://github.com/gitbrent/PptxGenJS/issues/562) ([Tehnix](https://github.com/Tehnix))
- Fixed: Exporting images corrupting file. [\#578](https://github.com/gitbrent/PptxGenJS/issues/578) ([joeberth](https://github.com/joeberth))
- Fixed: Multiple files getting downloaded if multiple base64 images are added. [\#581](https://github.com/gitbrent/PptxGenJS/issues/581) ([akshaymagapu](https://github.com/akshaymagapu))
- Fixed: Links in tables won't work on tables generated with autoPage [\#583](https://github.com/gitbrent/PptxGenJS/issues/583) ([githuis](https://github.com/githuis))
- Fixed: Added rounding of margin values to avoid invalid XML [\#633](https://github.com/gitbrent/PptxGenJS/pull/633) ([kevinresol](https://github.com/kevinresol))

### Removed

- Removed: jQuery is no longer required (!)

## [2.6.0] - 2019-09-24

### Added

- Host the Examples demo webpage online [\#505](https://github.com/gitbrent/PptxGenJS/pull/505) ([multiplegeorges](https://github.com/multiplegeorges))
- Add types key to package.json [\#529](https://github.com/gitbrent/PptxGenJS/pull/529) ([adamlong5](https://github.com/adamlong5))
- Add support for font family css when export HTML table to slide. [\#571](https://github.com/gitbrent/PptxGenJS/pull/571) ([Jank1310](https://github.com/twatson83))

### Changed

- Fixed: MIME type is ppt now instead of "application/zip"
- Fixed: Not Able to add background image from the www source [\#497](https://github.com/gitbrent/PptxGenJS/issues/497) ([nish25sp](https://github.com/nish25sp))
- Fixed: Set proper MIME type for PPTX presentation [\#471](https://github.com/gitbrent/PptxGenJS/issues/471) ([StefanBrand](https://github.com/StefanBrand))
- Fixed: lineDash Option is not in documentation [\#526](https://github.com/gitbrent/PptxGenJS/issues/526) ([Jank1310](https://github.com/Jank1310))
- Fixed: Downloading PPT in iOS using Safari does not work. File named as UNKNOWN. [\#540](https://github.com/gitbrent/PptxGenJS/issues/540) ([mustafagentrit](https://github.com/mustafagentrit))
- Fixed: ReferenceError: strXmlBullet is not defined [\#587](https://github.com/gitbrent/PptxGenJS/issues/587) ([Saurabh-Chandil](https://github.com/Saurabh-Chandil))
- Fixed: Getting paraPropXmlCore not defined error - line 4200 in pptxgen.bundle.js missing "var" declaration [\#596](https://github.com/gitbrent/PptxGenJS/issues/596) ([rajeearyal](https://github.com/rajeearyal))

### Removed

## [2.5.0] - 2019-02-08

### Added

- Make Shapes available for a front-end usage [\#137](https://github.com/gitbrent/PptxGenJS/issues/137) ([spamforhope](https://github.com/spamforhope))
- Ability to rotate chart axis labels (`catAxisLabelRotate`/`valAxisLabelRotate`) [\#378](https://github.com/gitbrent/PptxGenJS/issues/378) ([teejayvanslyke](https://github.com/teejayvanslyke))
- New Chart Type: 3D bar charts [\#384](https://github.com/gitbrent/PptxGenJS/pull/384) ([loictro](https://github.com/loictro))
- New Chart Feature: Add Data Labels to Scatter Charts [\#420](https://github.com/gitbrent/PptxGenJS/pull/420) ([ReimaFrgos](https://github.com/ReimaFrgos))
- Add new chart options: `catAxisLabelFontBold`,`dataLabelFontBold`,`legendFontFace`,`valAxisLabelFontBold` [\#426](https://github.com/gitbrent/PptxGenJS/issues/426) ([BandaSatish07](https://github.com/BandaSatish07))
- Add missing jpg content type to fix corrupt presentation for Office365 [\#435](https://github.com/gitbrent/PptxGenJS/pull/435) ([antonandreyev](https://github.com/antonandreyev))
- Add `catAxisMinVal` and `catAxisMaxVal` [\#462](https://github.com/gitbrent/PptxGenJS/pull/462) ([vrimar](https://github.com/vrimar))
- New Chart Option: `valAxisCrossesAt` [\#474](https://github.com/gitbrent/PptxGenJS/pull/474) ([ReimaFrgos](https://github.com/ReimaFrgos))
- Docs: Show how to save as Blob using client browser [\#478](https://github.com/gitbrent/PptxGenJS/issues/478) ([crazyx13th](https://github.com/crazyx13th))

### Changed

- Fixed: Dynamic Text Options do not apply [\#427](https://github.com/gitbrent/PptxGenJS/issues/427) ([sunnyar](https://github.com/sunnyar))
- Removed: legacy/deprecated attributes from README javascript script tags [\#431](https://github.com/gitbrent/PptxGenJS/pull/431) ([efx](https://github.com/efx))
- Fixed: issue with SlideNumber `fontSize` float values [\#432](https://github.com/gitbrent/PptxGenJS/issues/432) ([efx](https://github.com/efx))
- Fixed: query and fragment from image URL extension [\#433](https://github.com/gitbrent/PptxGenJS/pull/433) ([katsuya-horiuchi](https://github.com/katsuya-horiuchi))
- Changed: Replace "$" with "jQuery" to fix integration issues with some applications [\#436](https://github.com/gitbrent/PptxGenJS/pull/436) ([antonandreyev](https://github.com/antonandreyev))
- Changed: Export more types to enhance TypeScript support [\#443](https://github.com/gitbrent/PptxGenJS/pull/443) ([ntietz](https://github.com/ntietz))
- Fixed: Rounding in percentage leads to small deviations [\#470](https://github.com/gitbrent/PptxGenJS/pull/470) ([Slidemagic](https://github.com/Slidemagic)) [\#475](https://github.com/gitbrent/PptxGenJS/pull/475) ([ReimaFrgos](https://github.com/ReimaFrgos))
- Fixed: Hyperlinks causing duplicate relationship ID when other objects on page [\#477](https://github.com/gitbrent/PptxGenJS/pull/477) ([ReimaFrgos](https://github.com/ReimaFrgos))
- Fixed: ordering of paragraph properties [\#485](https://github.com/gitbrent/PptxGenJS/pull/485) ([sleepylemur](https://github.com/sleepylemur))

### Removed

## [2.4.0] - 2018-10-28

### Added

- Added support for SVG images [\#401](https://github.com/gitbrent/PptxGenJS/pull/401) ([Krelborn](https://github.com/Krelborn))
- Better detection/support for Angular [\#415](https://github.com/gitbrent/PptxGenJS/pull/415) ([antiremy](https://github.com/antiremy))

### Changed

- Demo page converted to Bootstrap 4 [gitbrent](https://github.com/gitbrent)
- Fixed issue with float font-sizes in `addSlidesForTable()` [gitbrent](https://github.com/gitbrent)
- No Color on negative bars when barGrouping is stacked [\#343](https://github.com/gitbrent/PptxGenJS/issues/343)
  ([vanarebane](https://github.com/vanarebane)) [\#419](https://github.com/gitbrent/PptxGenJS/pull/419)
  ([octy40](https://github.com/octy40))
- Improve typescript declaration files [\#409](https://github.com/gitbrent/PptxGenJS/pull/409) ([michaelbeaumont](https://github.com/michaelbeaumont))
- X and Y table coordinates with value of zero ignored [\#411](https://github.com/gitbrent/PptxGenJS/pull/411) ([tovab](https://github.com/tovab))
- Placeholder left align property needs fixing [\#417](https://github.com/gitbrent/PptxGenJS/pull/417) ([raphael-trzpit](https://github.com/raphael-trzpit))
- Replace jquery each by standard forEach [\#418](https://github.com/gitbrent/PptxGenJS/pull/418) ([fdussert](https://github.com/fdussert))
- BugFix: 0 value plot points ignored on Scatter Chart [\#422](https://github.com/gitbrent/PptxGenJS/pull/422) ([ReimaFrgos](https://github.com/ReimaFrgos))
- Pass the callback as a function, rather than invoke it [\#424](https://github.com/gitbrent/PptxGenJS/pull/424) ([danielsiwiec](https://github.com/danielsiwiec))

### Removed

## [v2.3.0](https://github.com/gitbrent/pptxgenjs/tree/v2.3.0) (2018-09-12)

[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v2.2.0...v2.3.0)

**Highlights:**

- New Feature: Placeholders
- New Feature: Speaker Notes
- `addImage()` can now load both local ("../img.png") and remote images ("<https://wikimedia.org/logo.jpg>")
- Typescript definitions are now available
- `jquery-node` replaced with latest `jquery` package [only affects npm users]

**Fixed Bugs:**

- Remove jquery-node dependency (fixes XSS Vulnerability Security Warning) [\#350](https://github.com/gitbrent/PptxGenJS/issues/350) ([TinkerJack](https://github.com/TinkerJack))
- Cannot set valAxisMinVal to 0 [\#357](https://github.com/gitbrent/PptxGenJS/issues/357) ([GiridharGNair](https://github.com/GiridharGNair))
- Multiple paragraph spacings if newline character occur in text [\#368](https://github.com/gitbrent/PptxGenJS/issues/368) ([vpetzel](https://github.com/vpetzel))
- Rotate working incorrectly [\#370](https://github.com/gitbrent/PptxGenJS/issues/370) ([michaelcbrook](https://github.com/michaelcbrook))
- Removed error thrown while rendering Multi Type chart containing Area [\#371](https://github.com/gitbrent/PptxGenJS/pull/371)
  ([KrishnaTejaReddyV](https://github.com/KrishnaTejaReddyV))
- Bugfix/enhancement for EncodeXML in speaker notes text [\#373](https://github.com/gitbrent/PptxGenJS/pull/373) ([travispwingo](https://github.com/travispwingo))

**Implemented Enhancements:**

- `addImage()` updated with new code allowing both local and remote images to be used (browser and Node). ([gitbrent](https://github.com/gitbrent))
- Typescript definitions have been created for the PptxGenJS API Methods (`pptxgen.d.ts`). ([gitbrent](https://github.com/gitbrent))
- New Feature: Placeholder support in Master Slides [\#359](https://github.com/gitbrent/PptxGenJS/pull/359) ([conbow](https://github.com/conbow))
- New Feature: Speaker Notes [\#239](https://github.com/gitbrent/PptxGenJS/issues/239) [\#361](https://github.com/gitbrent/PptxGenJS/pull/361) ([travispwingo](https://github.com/travispwingo))
- New Chart Option: `displayBlanksAs` [\#365](https://github.com/gitbrent/PptxGenJS/pull/365) ([guipas](https://github.com/guipas))
- New Feature: ability to hide slides [\#367](https://github.com/gitbrent/PptxGenJS/pull/367) ([ReimaFrgos](https://github.com/ReimaFrgos))
- Add second Cat Axis for Scatter and Bubble [\#372](https://github.com/gitbrent/PptxGenJS/pull/372) ([KrishnaTejaReddyV](https://github.com/KrishnaTejaReddyV))
- New Chart Type: Add radar chart implementation [\#386](https://github.com/gitbrent/PptxGenJS/pull/386) ([loictro](https://github.com/loictro))

## [v2.2.0](https://github.com/gitbrent/pptxgenjs/tree/v2.2.0) (2018-06-17)

[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v2.1.0...v2.2.0)

**Fixed Bugs:**

- Shapes: How to add vertical lines [\#272](https://github.com/gitbrent/PptxGenJS/issues/272) ([simonjcarr](https://github.com/simonjcarr))
- autoFit is missing 'Shrink text on overflow' variation? [\#330](https://github.com/gitbrent/PptxGenJS/issues/330) ([cdutson](https://github.com/cdutson))
- Rowspan, Colspan, and Multi-Row Headers Not Working [\#331](https://github.com/gitbrent/PptxGenJS/pull/331) ([skellman](https://github.com/skellman))([dwright-novetta](https://github.com/dwright-novetta))
- Isolate variables to the local scope [\#334](https://github.com/gitbrent/PptxGenJS/pull/334) ([edvinasbartkus](https://github.com/edvinasbartkus))
- `addMedia` of type='online' not working? [\#335](https://github.com/gitbrent/PptxGenJS/issues/335) ([lndev1](https://github.com/lndev1))
- Fixed Error thrown while rendering Area Chart [\#342](https://github.com/gitbrent/PptxGenJS/pull/342) ([KrishnaTejaReddyV](https://github.com/KrishnaTejaReddyV))
- Fixed Title display on showTitle = false error [\#344](https://github.com/gitbrent/PptxGenJS/pull/344) ([KrishnaTejaReddyV](https://github.com/KrishnaTejaReddyV))
- `getPageNumber()` is missing from the "Slide Methods" documentation [\#353](https://github.com/gitbrent/PptxGenJS/pull/353) ([kumaarraja](https://github.com/kumaarraja))

**Implemented Enhancements:**

- New Feature! `addImage()` and `addMedia()` methods now accept URLs [\#325](https://github.com/gitbrent/PptxGenJS/pull/325) ([gitbrent](https://github.com/gitbrent))
- Make Node detection more robust [\#277](https://github.com/gitbrent/PptxGenJS/issues/277) ([adrianirwin](https://github.com/adrianirwin)) ([DSheffield](https://github.com/DSheffield))
- Updated pptxgenjs-demo files to use CDNs instead of local files ([gitbrent](https://github.com/gitbrent))
- Updated Node.js detection to increase reliability for Angular users et al. ([gitbrent](https://github.com/gitbrent))
- Add `w` and `h` attributes to `slideNumber()` [\#336](https://github.com/gitbrent/PptxGenJS/issues/336) ([s7726](https://github.com/s7726))

## [v2.1.0](https://github.com/gitbrent/pptxgenjs/tree/v2.1.0) (2018-04-02)

[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v2.0.0...v2.1.0)

**Fixed Bugs:**

- HTML-to-PowerPoint is creating many extra columns with colspan [\#284](https://github.com/gitbrent/PptxGenJS/issues/284) ([svaak](https://github.com/svaak))
- HTML-to-PowerPoint rowspan is not working ([gitbrent](https://github.com/gitbrent))
- Fix docs/examples to use new fontSize, remove unsupported font_size [\#297](https://github.com/gitbrent/PptxGenJS/issues/297) ([pstoll](https://github.com/pstoll))

**Implemented Enhancements:**

- Mis-detecting Existence of Node.js [\#277](https://github.com/gitbrent/PptxGenJS/issues/277) ([adrianirwin](https://github.com/adrianirwin)) ([DSheffield](https://github.com/DSheffield))
- Add Text Outline functionality [\#298](https://github.com/gitbrent/PptxGenJS/issues/298) ([stevenljacobsen](https://github.com/stevenljacobsen))
- Adding rounded corners to images [\#309](https://github.com/gitbrent/PptxGenJS/issues/309) ([hoangpq](https://github.com/hoangpq))

## [v2.0.0](https://github.com/gitbrent/pptxgenjs/tree/v2.0.0) (2018-01-23)

[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.10.0...v2.0.0)

**BREAKING CHANGES**

- NodeJS instantiation is now standard (see Issue [\#83](https://github.com/gitbrent/PptxGenJS/issues/83) and `examples/nodejs-demo.js`) which now allows new instances/presentations
- (See "Version 2.0 Breaking Changes" in the README for a complete list)

**Fixed Bugs:**

- Master Slide slide number doesn't show using "New Slide" PPT Function [\#229](https://github.com/gitbrent/PptxGenJS/issues/229) ([ineran](https://github.com/ineran))
- Values of 0 (zero) in series are missing in line chart [\#240](https://github.com/gitbrent/PptxGenJS/issues/240) ([andrei-cs](https://github.com/andrei-cs))
- Node: "DeprecationWarning: Calling an asynchronous function without callback is deprecated." [\#252](https://github.com/gitbrent/PptxGenJS/issues/252) ([the-yadu](https://github.com/the-yadu))
- The UP_DOWN_ARROW shape appears to have duplicate keys [\#253](https://github.com/gitbrent/PptxGenJS/issues/253) ([heavysixer](https://github.com/heavysixer))
- Local demo can not run in IE [\#273](https://github.com/gitbrent/PptxGenJS/issues/273) ([IvanTao](https://github.com/IvanTao))

**Implemented Enhancements:**

- Is it possible to link from one slide to another? [\#251](https://github.com/gitbrent/PptxGenJS/issues/251) ([heavysixer](https://github.com/heavysixer))
- Add rot and vert options to text body properties [\#254](https://github.com/gitbrent/PptxGenJS/issues/254) ([level46](https://github.com/level46))
- Add Character Spacing option [\#265](https://github.com/gitbrent/PptxGenJS/issues/265) ([nguyenhuuphuc83](https://github.com/nguyenhuuphuc83))

## [v1.10.0](https://github.com/gitbrent/pptxgenjs/tree/v1.10.0) (2017-11-14)

[Full Changelog](https://github.com/gitbrent/pptxgenjs/compare/v1.9.0...v1.10.0)

**Fixed Bugs:**

- Fixed bug that was preventing 'chartColorsOpacity' from being anything other than 50 percent. ([gitbrent](https://github.com/gitbrent))
- The `newPageStartY` option is not being honored by `addSlidesForTable()` [\#222](https://github.com/gitbrent/PptxGenJS/issues/222) ([shaunvdp](https://github.com/shaunvdp))
- Line chart with one series displays broken [\#225](https://github.com/gitbrent/PptxGenJS/issues/225) ([andrei-cs](https://github.com/andrei-cs))
- The `*AxisLineShow` chart options do not work [\#231](https://github.com/gitbrent/PptxGenJS/pull/231) ([mconlin](https://github.com/mconlin))

**Implemented Enhancements:**

- New chart type: bubble charts [\#208](https://github.com/gitbrent/PptxGenJS/issues/208) ([shrikantbhongade](https://github.com/shrikantbhongade))
- New Chart option: Legend Text Color [\#233](https://github.com/gitbrent/PptxGenJS/issues/233) ([mconlin](https://github.com/mconlin))
- New Text option: `strike` [\#238](https://github.com/gitbrent/PptxGenJS/issues/238) ([adrienco88](https://github.com/adrienco88))

## [v1.9.0](https://github.com/gitbrent/pptxgenjs/tree/v1.9.0) (2017-10-10)

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
