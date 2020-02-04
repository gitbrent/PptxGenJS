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
2. Update `src/pptxgen.ts` version
3. Build using `$ npm run ship`
4. Consolidate new changes from `src/bld/*.ts` into `types/index.d.ts`
5. Open `dist/*.js` and check headers
6. Update `CHANGELOG.md` with new date
7. Update `README.md` with new CDN links

## Test Newest Library Build

### Browser

Run all tests in browser

- [Local Demo](file:///Users/brentely/GitHub/PptxGenJS/demos/browser/index.html)

### Node

Run Node test

```bash
$ cd ~/GitHub/PptxGenJS/demos/node
$ node demo.js All
```

### React/TypeScript

React Test

1. Ensure newest `dist/pptxgen.es.js` and `types/index.d.ts` under local node_modules
2. Update `demos/react-demo/package.json` version
3. Open `demos/react-demo/src/latest/Test.tsx`
4. Check existing code
5. Test defs by using auto-complete, "pptxgen.ChartType." etc.

```bash
$ cd ~/GitHub/PptxGenJS/demos/react-demo
$ npm run start
```

1. Go to http://localhost:3000 on iMac
2. Run both demo tests
3. Go to http://192.168.x.x:3000 on iPhone
4. Run both demo tests
5. Ensure each is viewable upon download
6. `npm run build`
7. copy entire "build" folder to Downloads for subsequently updating gh-pages with latest build (DO NOT use the deploy script offered onscreen!)

**NOTE** Any updates to `node_modules/dist/pptxgen.es.js` are not picked up by the server (ctrl-C and restart)

## Release New Version

### Pre-Release Check

1. Revert scripts in `./demos/browser/index.html`
2. Is version updated in package.json and pptxgen.js?
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

1. Save output from all tests and html2ppt for this release
2. Go test CDN links on README
3. Load **gh-pages** branch
4. Update `installation.md` with latest CDN version
5. Update demo-react by copying contents of the newest "build" (from above) into `./demo-react` folder
6. Update other documentation as needed
