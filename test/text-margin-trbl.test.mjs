/**
 * Text margin arrays are documented as TRBL [top, right, bottom, left].
 * Regression: textObjectToXml previously mapped them as LRBT (lIns←[0], tIns←[3]).
 */
import { describe, it } from 'node:test'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import PptxGenJS from '../dist/pptxgen.cjs.js'

const ONEPT = 12700

function ptsToEmu(pt) {
	return Math.round(pt * ONEPT)
}

async function bodyPrAttrsFromPptx(buffer) {
	const zip = await JSZip.loadAsync(buffer)
	const slideXml = await zip.file('ppt/slides/slide1.xml').async('string')
	const bodyPr = slideXml.match(/<a:bodyPr\b[^>]*>/)
	assert.ok(bodyPr, 'expected <a:bodyPr> in slide1.xml')
	const attrs = Object.fromEntries(
		[...bodyPr[0].matchAll(/\b(lIns|tIns|rIns|bIns)="(\d+)"/g)].map(([, k, v]) => [k, Number(v)])
	)
	return attrs
}

describe('text margin TRBL mapping', () => {
	it('maps margin [top, right, bottom, left] to tIns/rIns/bIns/lIns', async () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()
		// Distinct sides so a swap would fail loudly
		slide.addText('bullet list', {
			x: 0.5,
			y: 0.5,
			w: 4,
			h: 2,
			margin: [10, 20, 30, 40],
		})

		const buffer = await pptx.write({ outputType: 'nodebuffer' })
		const insets = await bodyPrAttrsFromPptx(buffer)

		assert.deepEqual(insets, {
			tIns: ptsToEmu(10),
			rIns: ptsToEmu(20),
			bIns: ptsToEmu(30),
			lIns: ptsToEmu(40),
		})
	})

	it('applies a scalar margin to all four sides', async () => {
		const pptx = new PptxGenJS()
		const slide = pptx.addSlide()
		slide.addText('uniform', {
			x: 0.5,
			y: 0.5,
			w: 4,
			h: 1,
			margin: 7,
		})

		const buffer = await pptx.write({ outputType: 'nodebuffer' })
		const insets = await bodyPrAttrsFromPptx(buffer)
		const expected = ptsToEmu(7)

		assert.deepEqual(insets, {
			tIns: expected,
			rIns: expected,
			bIns: expected,
			lIns: expected,
		})
	})
})
