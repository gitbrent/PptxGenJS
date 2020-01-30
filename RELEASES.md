# PptxGenJS Release Checklist

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->


- [Test Library](#test-library)
  - [Browser](#browser)
  - [Node](#node)
  - [React/TypeScript](#reacttypescript)
- [Release New Version](#release-new-version)
  - [File Prep](#file-prep)
  - [GitHub](#github)
  - [NPM](#npm)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

## Test Library

### Browser

Enable watch/continuous build

```bash
$ cd ~/GitHub/PptxGenJS
$ gulp
```

Run all tests in browser using: [Local Demo](file:///Users/brentely/GitHub/PptxGenJS/demos/browser/index.html)

### Node

Run Node test

```bash
$ cd ~/GitHub/PptxGenJS/demos/node
$ node demo.js All -local
```

### React/TypeScript

React Test

```bash
$ cd ~/GitHub/PptxGenJS/demos/react-demo
$ npm run start
```

- Go to https://localhost:3000 and run the demo

TypeScript Defs

- Open `demos/react/demo/src/latest/Test.tsx`
- Check existing code
- Test defs by using auto-complete, "pptxgen.chart" etc.
- Note: Copy newest `types/index.d.ts` to local node_modules if updated recently

## Release New Version

### File Prep
1. Consolidate `src/bld/*.ts` into `types/index.d.ts` (Note: `charts` and `shapes` are special and stay!)
2. Update `package.json` version
3. Update `src/pptxgen.ts` version
4. Build using `$ gulp`
5. Open `dist/*.js` and check headers
6. Update `CHANGELOG.md` date, etc.
7. Checks: Are ts-defs updated?

### GitHub
1. Copy CHANGELOG entry and draft new release: [Releases](https://github.com/gitbrent/PptxGenJS/releases)
2. Use "Version X.x.x" as title and "v3.1.1" as tag
3. Go back to Releases page, double-check title/tag, release when ready

### NPM
1. `cd ~/GitHub/PptxGenJS`
2. `npm publish`
