---
id: usage-add-slide
title: Adding a Slide
---

## Syntax

```javascript
let slide = pptx.addSlide();
```

## Slide Templates/Master Slides

### Master Slide Syntax

```javascript
let slide = pptx.addSlide("MASTER_NAME");
```

(See [Master Slides](/PptxGenJS/docs/masters) for more about creating masters/templates)

### Master Slide Examples

```javascript
// Create a new Slide that will inherit properties from a pre-defined master page (margins, logos, text, background, etc.)
let slide = pptx.addSlide("TITLE_SLIDE");
```

## Slides Return Themselves

The Slide object returns a reference to itself, so calls can be chained.

Example:

```javascript
slide.addImage({ path: "img1.png", x: 1, y: 2 }).addImage({ path: "img2.jpg", x: 5, y: 3 });
```

## Slide Methods
See [Slide Methods](/PptxGenJS/docs/usage-slide-options) for features such as Background and Slide Numbers.
