---
id: needs-repair-errors
title: PowerPoint "Needs Repair" Errors
sidebar_label: Needs Repair Errors
---

Troubleshooting guide for when you're encountering an error where PowerPoint shows a "needs repair" error dialog when opening your generated presentation.

## Why these errors are difficult to debug

PowerPoint's files are essentially ZIP archives containing many XML files. When PowerPoint reports a "needs repair" error, it means that one or more of these internal XML files do not conform to the Office Open XML (OOXML) specification.

Unfortunately, PowerPoint's error messages are generic and do not pinpoint the exact line, element, or file within the XML that is causing the problem. It's like finding a single syntax error in a massive, undocumented codebase without a debugger. It happends to me all the time when adding new features to the library, and it super sucks!

## How to diagnose your specific issue

Since `pptxgenjs` generates the OOXML based on your API calls, the most effective method for identifying the root cause of **your particular error** is a process of elimination:

### Isolate the problematic slide

1. Start by generating your presentation with **only a few slides**, or even just one.
2. If that works, gradually **add slides back one by one**. Generate and open the `.pptx` file after each addition.
3. The moment the presentation becomes unreadable, you've identified the slide that contains the problematic content.

### Pinpoint the problematic feature

Once you've isolated the problematic slide, begin removing content from that specific slide.

- Remove elements such as:
  - Textboxes
  - Images
  - Tables
  - Charts
  - Shapes
- Remove these features **one by one**, generating and testing the file after each removal. This will help you narrow down which specific feature or combination of features is causing the XML validation error.

Alternatuvely, try different options on auto-paged tables, charts, etc. It's often the case that bad/incorrect options cause errors.

## What to do once you've found the cause

### Review your API usage/options

Double-check the options and data you are passing for the identified problematic feature against the `pptxgenjs` documentation. Minor typos, incorrect data types, or out-of-bounds values can easily lead to invalid XML.

Remember, there are working code examples for every available option. Start with code that works, then modify from there.

### Search for existing issues

Check the `pptxgenjs` GitHub issues (both open and closed) for the specific feature you've identified. Someone else might have encountered and reported a similar problem.

### Open a new, detailed issue

If you're confident you've found a bug in `pptxgenjs`, please open a **new GitHub issue**. In your report, **be sure to include:**

- The `pptxgenjs` version you are using.
- A **minimal reproducible code example** that demonstrates the issue (only the problematic slide/feature).
- Any relevant error messages from your browser console or Node.js environment.

Your detailed investigation helps us immensely in identifying and fixing bugs in the library.
