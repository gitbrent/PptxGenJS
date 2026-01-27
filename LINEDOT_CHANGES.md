# LineDot Changes to @linedotai/js-to-ppt

This document tracks all modifications made to the forked pptxgenjs library for LineDot's PowerPoint generation needs.

## Fork Information

- **Original Library**: [PptxGenJS](https://github.com/gitbrent/PptxGenJS) by Brent Ely
- **Fork Name**: `@linedotai/js-to-ppt`
- **Base Version**: v4.0.1 (forked from upstream)
- **Current Version**: 1.0.5

---

## Version History

### v1.0.5 (Unreleased)

**Bug Fix: Bold Inline Text Causing Line Breaks**

#### Problem
When text contained inline bold formatting (e.g., "We employ **FFD bin packing** to concatenate"), PowerPoint would render unwanted line breaks after bold spans. The text would break like:
```
We employ FFD bin packing
to concatenate input vectors
```
Instead of rendering as a continuous line.

#### Root Cause
In `src/gen-xml.ts`, the `genXmlTextBody()` function was generating `<a:pPr>` (paragraph properties) elements for **every text run** within a paragraph. According to OOXML specification, `<a:pPr>` should only appear **once** per `<a:p>` element, immediately after the opening tag.

When multiple `<a:pPr>` elements were present (one for normal text, one for bold text), PowerPoint interpreted each as starting a new paragraph-like structure, causing visual line breaks.

#### Fix
Modified `src/gen-xml.ts` (around line 1377) to only generate `<a:pPr>` for the first text run (`idx === 0`) in each paragraph:

```typescript
// BEFORE (buggy):
paragraphPropXml = genXmlParagraphProperties(textObj, false)
strSlideXml += paragraphPropXml.replace('<a:pPr></a:pPr>', '')

// AFTER (fixed):
// IMPORTANT: Only add paragraph properties (<a:pPr>) for the FIRST text run in a paragraph.
// In OOXML, <a:pPr> should only appear once per <a:p>, right after the opening tag.
// Adding it for every text run causes rendering issues (like unwanted line breaks after bold text).
if (idx === 0) {
    paragraphPropXml = genXmlParagraphProperties(textObj, false)
    strSlideXml += paragraphPropXml.replace('<a:pPr></a:pPr>', '')
}
```

#### Files Changed
- `src/gen-xml.ts` - Added `idx === 0` check for paragraph properties
- `dist/pptxgen.cjs.js` - Rebuilt bundle
- `dist/pptxgen.es.js` - Rebuilt bundle
- `package.json` - Version bump to 1.0.5

---

### v1.0.4

- Version alignment release

### v1.0.3

- Shrink-to-fit functionality improvements

### v1.0.2

- Chart rendering fix for Keynote compatibility

### v1.0.1

- Initial LineDot fork from pptxgenjs v4.0.1
- Package renamed to `@linedotai/js-to-ppt`
- Published to GitHub Package Registry

---

## Consumer-Side Companion Fix

Note: The library fix alone may not be sufficient. The consuming code in `linedot-backend` must also ensure consistent paragraph properties across all text spans in a block.

In `linedot-backend/src/ppt/utils/pptx-generator.ts`, ensure the `align` property is set consistently on **all spans** within the same block, not just the first one:

```typescript
// Set align on ALL spans to prevent pptxGenJS from detecting property changes
spans.forEach(span => {
  if (!span.options.align && block.align) {
    span.options.align = block.align;
  }
});
```

This prevents pptxGenJS from splitting text runs when it detects differing properties between spans.

---

## Build Instructions

After making changes to source files:

```bash
cd js-to-ppt
npm run build
```

To use in linedot-backend during development:

```bash
cd js-to-ppt
npm link

cd ../linedot-backend
npm link @linedotai/js-to-ppt
```

---

## Related Issues

- Internal: Bold text line break issue in PPTX generation
- OOXML Spec: `<a:pPr>` should only appear once per `<a:p>` element
