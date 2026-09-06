/**
 * ECMA-376: p:cNvPr/@id must be unique within a slide's p:spTree.
 * Regression for gitbrent/PptxGenJS#1532 (table + other shape).
 */
const { test } = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const JSZip = require('jszip')

function loadPptxGenJS () {
	const built = path.join(__dirname, '../src/bld/pptxgen.cjs.js')
	const dist = path.join(__dirname, '../dist/pptxgen.cjs.js')
	return require(fs.existsSync(built) ? built : dist)
}

function extractCnvPrIds (xml) {
	return [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map(m => Number(m[1]))
}

async function slideXml (pptx, slidePath) {
	const buf = await pptx.write({ outputType: 'nodebuffer' })
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file(slidePath)
	assert.ok(entry, `missing zip entry ${slidePath}`)
	return entry.async('string')
}

function assertUniqueCnvPrIds (ids, label) {
	assert.ok(ids.length >= 3, `${label}: expected nvGrpSpPr + objects, got [${ids.join(', ')}]`)
	assert.equal(ids[0], 1, `${label}: p:nvGrpSpPr must keep id="1"`)
	assert.equal(
		new Set(ids).size,
		ids.length,
		`${label}: duplicate p:cNvPr/@id [${ids.join(', ')}]`
	)
}

test('p:cNvPr/@id values are unique when a slide has a table and a text shape', async () => {
	const PptxGenJS = loadPptxGenJS()
	const pptx = new PptxGenJS()
	const slide = pptx.addSlide()
	slide.addText('A text box', { x: 0.5, y: 0.5, w: 4, h: 0.5 })
	slide.addTable([[{ text: 'A' }, { text: 'B' }]], { x: 0.5, y: 1.5, w: 4 })

	const xml = await slideXml(pptx, 'ppt/slides/slide1.xml')
	assertUniqueCnvPrIds(extractCnvPrIds(xml), 'slide1 text+table')
})

test('p:cNvPr/@id values stay unique on later slides and with mixed object order', async () => {
	const PptxGenJS = loadPptxGenJS()
	const pptx = new PptxGenJS()
	pptx.addSlide()
	const slide = pptx.addSlide()
	slide.addTable([[{ text: 'A' }, { text: 'B' }]], { x: 0.5, y: 1.5, w: 4 })
	slide.addText('A text box', { x: 0.5, y: 0.5, w: 4, h: 0.5 })
	slide.addShape(pptx.ShapeType.rect, { x: 5, y: 0.5, w: 1, h: 1 })

	const xml = await slideXml(pptx, 'ppt/slides/slide2.xml')
	assertUniqueCnvPrIds(extractCnvPrIds(xml), 'slide2 table+text+shape')
})
