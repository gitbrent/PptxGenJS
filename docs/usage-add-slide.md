---
id: usage-add-slide
title: Adding a Slide
---

## Syntax

Create a new slide in the presentation:

```javascript
let slide = pptx.addSlide();
```

## Returns

The `addSlide()` method returns a reference to the created Slide object, so method calls can be chained.

```javascript
let slide1 = pptx.addSlide();
slide1
  .addImage({ path: "img1.png", x: 1, y: 2 })
  .addImage({ path: "img2.jpg", x: 5, y: 3 });
```

You can also create multiple slides:

```javascript
let slide1 = pptx.addSlide();
slide1.addText("Slide One", { x: 1, y: 1 });

let slide2 = pptx.addSlide();
slide2.addText("Slide Two", { x: 1, y: 1 });
```

## Slide Methods

See [Slide Methods](/PptxGenJS/docs/usage-slide-options) for features such as Background and Slide Numbers.

## Slide Masters

Want to use a layout with predefined logos, margins, or styles? See [Slide Masters](/PptxGenJS/docs/masters) to learn how to create and apply slide masters.
