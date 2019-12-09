import { TEXT_VALIGN, TEXT_HALIGN } from '../core-enums'

import Bullet from './bullet'

const alignment = align => {
	switch (align) {
		case 'left':
			return ' algn="l"'
		case 'right':
			return ' algn="r"'
		case 'center':
			return ' algn="ctr"'
		case 'justify':
			return ' algn="just"'
		default:
			return ''
	}
}

export default class TextElement {
	bullet
	align
	lineSpacing
	indentLevel
	paraSpaceBefore
	paraSpaceAfter
	rtlMode

	constructor({ rtlMode, paraSpaceBefore, paraSpaceAfter, indentLevel, bullet, align, lineSpacing }) {
		if (indentLevel && !isNaN(Number(indentLevel)) && Number(this.indentLevel) > 0) {
			this.indentLevel = Number(indentLevel)
		}
		if (paraSpaceBefore && !isNaN(Number(paraSpaceBefore)) && Number(this.paraSpaceBefore) > 0) {
			this.paraSpaceBefore = Number(paraSpaceBefore)
		}
		if (paraSpaceAfter && !isNaN(Number(paraSpaceAfter)) && Number(this.paraSpaceAfter) > 0) {
			this.paraSpaceAfter = Number(paraSpaceAfter)
		}

		this.bullet = new Bullet(bullet)

		const alignInput = (align || '').toLowerCase()
		if (alignInput.startsWith('c')) this.align = TEXT_HALIGN.center
		else if (alignInput.startsWith('l')) this.align = TEXT_HALIGN.left
		else if (alignInput.startsWith('r')) this.align = TEXT_HALIGN.right
		else if (alignInput.startsWith('j')) this.align = TEXT_HALIGN.justify

		this.rtlMode = rtlMode

		this.lineSpacing = lineSpacing && !isNaN(lineSpacing) ? lineSpacing : null
	}

	render(tag, body = '') {
		let bulletLvl0Margin = 342900

		const marginLeft = this.indentLevel && this.indentLevel > 0 ? bulletLvl0Margin + bulletLvl0Margin * this.indentLevel : bulletLvl0Margin

		return `
        <${tag} ${[
			this.rtlMode ? ' rtl="1" ' : '',
			alignment(this.align),
			this.indentLevel ? ` lvl="${this.indentLevel}"` : '',
			this.bullet.enabled ? ` marL="${marginLeft}" indent="-${bulletLvl0Margin}"` : '',
		].join('')}>
          ${[
				// IMPORTANT: the body element require strict ordering - anything out of order is ignored. (PPT-Online, PPT for Mac)
				this.lineSpacing ? `<a:lnSpc><a:spcPts val="${this.lineSpacing}00"/></a:lnSpc>` : '',
				this.paraSpaceBefore ? `<a:spcBef><a:spcPts val="${this.paraSpaceBefore * 100}"/></a:spcBef>` : '',
				this.paraSpaceAfter ? `<a:spcAft><a:spcPts val="${this.paraSpaceAfter * 100}"/></a:spcAft>` : '',
				this.bullet.render(),
			].join('')}
			${body}
        </${tag}>
    `
	}
}
