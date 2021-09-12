---
id: sections
title: Slide Sections
---

Group slides using sections.

## Syntax

```typescript
pptx.addSection({ title: "Tables" });
pptx.addSection({ title: "Tables", order: 3 });
```

## Section Options

| Option  | Type    | Description   | Possible Values                                                             |
| :------ | :------ | :------------ | :-------------------------------------------------------------------------- |
| `title` | string  | section title | 0-n OR 'n%'. (Ex: `{x:'50%'}` will place object in the middle of the Slide) |
| `order` | integer | section order | 1-n. Used to add section at any index                                       |

## Section Example

```typescript
import pptxgen from "pptxgenjs";
let pptx = new pptxgen();

// STEP 1: Create a section
pptx.addSection({ title: "Tables" });

// STEP 2: Provide section title to a slide that you want in corresponding section
let slide = pptx.addSlide({ sectionTitle: "Tables" });

slide.addText("This slide is in the Tables section!", { x: 1.5, y: 1.5, fontSize: 18, color: "363636" });
pptx.writeFile({ fileName: "Section Sample.pptx" });
```
