import { getSmartParseNumber, convertRotationDegrees } from '../gen-utils'

export default class Position {
	x
	y
	w
	h

	flipH
	flipV
	rotate

	constructor({ x, y, w, h, flipH, flipV, rotate }) {
		this.x = x
		this.y = y
		this.w = w
		this.h = h

		this.flipH = flipH
		this.flipV = flipV
		this.rotate = rotate
	}

	cx(presLayout) {
		if (typeof this.w !== 'undefined') return getSmartParseNumber(this.w, 'X', presLayout)
	}

	cy(presLayout) {
		if (typeof this.h !== 'undefined') return getSmartParseNumber(this.h, 'Y', presLayout)
	}

	render(presLayout) {
		if (typeof this.x === 'undefined' && typeof this.y === 'undefined' && typeof this.w === 'undefined' && typeof this.h === 'undefined') {
			return ''
		}

		let locationAttr = ''
		let x
		let y
		let cx
		let cy

		if (typeof this.x !== 'undefined') x = getSmartParseNumber(this.x, 'X', presLayout)
		if (typeof this.y !== 'undefined') y = getSmartParseNumber(this.y, 'Y', presLayout)
		if (typeof this.w !== 'undefined') cx = getSmartParseNumber(this.w, 'X', presLayout)
		if (typeof this.h !== 'undefined') cy = getSmartParseNumber(this.h, 'Y', presLayout)

		if (this.flipH) locationAttr += ' flipH="1"'
		if (this.flipV) locationAttr += ' flipV="1"'
		if (this.rotate) locationAttr += ' rot="' + convertRotationDegrees(this.rotate) + '"'

		return `
            <a:xfrm${locationAttr}>
                <a:off x="${x}" y="${y}"/>
                <a:ext cx="${cx}" cy="${cy}"/>
            </a:xfrm>`
	}
}
