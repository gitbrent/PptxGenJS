# PptxGenJS Release Checklist

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

1. Consolidate `src/bld/*.ts` into `types/index.d.ts` (Note: `charts` and `shapes` are special and stay!)
2. Update `package.json` version
3. Update `src/pptxgen.ts` version
4. Build using `$ gulp`
5. Open `dist/*.js` and check headers
6. Update `CHANGELOG.md` date, etc.
7. Copy CHANGELOG entry and draft new release: [Releases](https://github.com/gitbrent/PptxGenJS/releases)
8. Use "Version X.x.x" as title and "v3.1.1" as tag
9. Go back to Releases page, double-check title/tag, release when ready
10. `cd ~/GitHub/PptxGenJS` then `npm publish`
