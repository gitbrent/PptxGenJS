---
id: deprecated
title: Deprecated
---

## Version 3.0 Breaking Changes

Please see the [Version 3.0 Migration Guide](https://github.com/gitbrent/PptxGenJS/wiki/Version-3.0-Migration-Guide)


## Version 2.0 Breaking Changes

Please note that version 2.0.0 enabled some much needed cleanup, but may break your previous code...
(however, a quick search-and-replace will fix any issues).

While the changes may only impact cosmetic properties, it's recommended you test your solutions thoroughly before upgrading PptxGenJS to the 2.0 version.

### All Users
The library `getVersion()` method is now a property: `version`

Option names are now caseCase across all methods:
* `font_face` renamed to `fontFace`
* `font_size` renamed to `fontSize`
* `line_dash` renamed to `lineDash`
* `line_head` renamed to `lineHead`
* `line_size` renamed to `lineSize`
* `line_tail` renamed to `lineTail`

Options deprecated in early 1.0 versions (hopefully nobody still uses these):
* `marginPt` renamed to `margin`


### Node Users
**Major Change**
* `require('pptxgenjs')` no longer returns a singleton instance
* `pptx = new PptxGenJS()` will create a single, unique instance
* Advantage: Creating [multiple presentations](#saving-multiple-presentations) is much easier now - see [Issue #83](https://github.com/gitbrent/PptxGenJS/issues/83) for more).
