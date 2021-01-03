# PptxGenJS Release Checklist

<!-- START doctoc generated TOC please keep comment here to allow auto update -->
<!-- DON'T EDIT THIS SECTION, INSTEAD RE-RUN doctoc TO UPDATE -->

- [Build Library, Update Files](#build-library-update-files)
- [Test Newest Library Build](#test-newest-library-build)
  - [Browser](#browser)
  - [Node](#node)
  - [React/TypeScript](#reacttypescript)
- [Release New Version](#release-new-version)
  - [Pre-Release Check](#pre-release-check)
  - [GitHub](#github)
  - [NPM](#npm)
- [Post-Release Tasks](#post-release-tasks)

<!-- END doctoc generated TOC please keep comment here to allow auto update -->

## Build Library, Update Files

1. Update `package.json` version
2. Update `src/pptxgen.ts` version (eg: `const VERSION = '3.3.1'`)
3. Update `CHANGELOG.md` with new date
4. Update `README.md` with new CDN links
5. Build library: npm scripts > `ship`
6. Consolidate new changes from `src/bld/*.ts` into `types/index.d.ts` and update version in head comment
7. Open `dist/*.js` and check headers

## Run Platform Tests

### Browser Test

1. Ensure newest `pptxgen.bundle.js` is loaded using F12 > Sources tab
2. Run all tests in browser [Demo Page](file:///Users/brentely/GitHub/PptxGenJS/demos/browser/index.html)

### Node Test

1. Update `demos/node/package.json` version
2. Run various tests

```bash
$ cd ~/GitHub/PptxGenJS/demos/node
$ npm run demo
$ npm run demo-all
$ npm run demo-text
$ npm run demo-stream
```

### React/TypeScript Test

1. Ensure newest `dist/pptxgen.es.js` and `types/index.d.ts` under local node_modules
2. Update `demos/react-demo/package.json` version
3. Open `demos/react-demo/src/tstest/Test.tsx`
4. Check existing code
5. Test defs by using auto-complete, "pptxgen.ChartType." etc.

```bash
$ cd ~/GitHub/PptxGenJS/demos/react-demo
$ npm run start
```

1. Go to [React Test](http://localhost:3000) on iMac, run demo tests
2. Go to http://192.168.1.x:3000 on iPhone, run demo tests
3. Go to http://192.168.1.x:3000 on Android, run demo tests
4. Open exports on each device to ensure MIME type is correct, looks right, etc.

```bash
$ cd ~/GitHub/PptxGenJS/demos/react-demo
$ npm run build
```

1. Copy entire "build" folder to Downloads for subsequently updating gh-pages with latest build (DO NOT use the deploy script offered onscreen!)

**NOTE** Any updates to `node_modules/dist/pptxgen.es.js` are not picked up by the server (ctrl-C and restart)

## Release New Version

### Pre-Release Check

1. Update `demos/browser/index.html` version and CDN links
2. Is version updated in package.json and src/pptxgen.ts?
3. Are `index.d.ts` defs updated?

### GitHub

1. Checkin all changes via GitHub Desktop
2. Copy CHANGELOG entry and draft new release: [Releases](https://github.com/gitbrent/PptxGenJS/releases)
3. Use "Version X.x.x" as title and "v3.1.1" as tag
4. Go back to Releases page, double-check title/tag, release when ready

### NPM

1. `cd ~/GitHub/PptxGenJS`
2. `npm publish`

## Post-Release Tasks

1. Go test CDN links on README
2. Load **gh-pages** branch
3. Update `installation.md` with latest CDN version
4. Update demo-react by copying contents of the newest "build" (from above) into `./demo-react` folder
5. Update `demo/index.html` with newest release
6. Update API documentation as needed
