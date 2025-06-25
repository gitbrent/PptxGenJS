# PptxGenJS Release Checklist

> This guide documents how to perform a PptxGenJS release.
> Maintainers should follow this checklist before pushing to npm or GitHub.

## 📋 Beta Releases

1. Update `package.json` version (ex: `4.1.0-beta.0`)
2. Update `src/pptxgen.ts` version
3. Build library: npm scripts > `ship`
4. `npm publish --tag beta`

## 🚀 Build Library, Update Files

1. Update `package.json` version
2. Update `src/pptxgen.ts` version (eg: `const VERSION = '4.0.1'`)
3. Update `CHANGELOG.md` with new date
4. Build library: npm scripts > `ship`
5. Consolidate new changes from `src/bld/*.ts` into `types/index.d.ts` and update version in head comment
6. Open `dist/*.js` and check headers
7. Update version in: `demos/node/package.json`
8. Update pptxgenjs dep version in: `demos/vite-demo/package.json`

## 🧪 Run Tests Before Release

### ⚠️ Run Standard Test Suite

See [TESTING.md](./TESTING.md) for complete test instructions.

### ⚠️ Capture Testing Results

| Dist File         | Test       | Tested Via             | Result |
| ----------------- | ---------- | ---------------------- | ------ |
| pptxgen.es.js     | Webpack 4  | SPFx (v1.16.1) project | ✅?🟡    |
| pptxgen.es.js     | Webpack 5  | SPFx (v1.19.1) project | ✅?🟡    |
| pptxgen.es.js     | Rollup 4   | Vite (v6) demo         | ✅?🟡    |
| pptxgen.cjs.js    | Node/CJS   | Node demo              | ✅?🟡    |
| pptxgen.bundle.js | Script     | Browser demo (desktop) | ✅?🟡    |
| pptxgen.bundle.js | Script     | Browser demo (iOS)     | ✅?🟡    |
| pptxgen.bundle.js | Web Worker | worker_test demo       | ✅?🟡    |

## 🚌 Release New Version

### 🟡 Pre-Release Checklist

1. Update: `demos/browser/index.html` head to use "RELEASE (CDN)"
2. Check: Is `version` updated in package.json?
3. Check: Is `version` updated in src/pptxgen.ts?
4. Check: Is `types/index.d.ts` version in header updated?

### 🟢 Release: GitHub

1. Checkin all changes via GitHub Desktop
2. Merge working branch into `main`
3. Copy CHANGELOG entry and draft new release: [Releases](https://github.com/gitbrent/PptxGenJS/releases)
4. Use "Version x.x.x" as title and "vX.X.X" as tag
5. Go back to Releases page, double-check title/tag, release when ready

### 🟢 Release: NPM

```bash
cd ~/GitHub/PptxGenJS
npm publish
```

## 🏁 Post-Release Tasks

1. Test CDN links on README.md
2. Load **gh-pages** branch
3. Update `installation.md` with latest CDN version
4. Copy contents of the newest "build" folder (from above) into `./demo-react` folder
5. Update API documentation if needed
