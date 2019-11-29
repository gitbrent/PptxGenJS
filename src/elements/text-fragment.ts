import {
	BULLET_TYPES,
	CRLF,
	DEF_CELL_BORDER,
	DEF_CELL_MARGIN_PT,
	EMU,
	LAYOUT_IDX_SERIES_BASE,
	ONEPT,
	PLACEHOLDER_TYPES,
	SLDNUMFLDID,
	SLIDE_OBJECT_TYPES,
	DEF_PRES_LAYOUT_NAME,
	TEXT_HALIGN,
} from '../core-enums'
import { PowerPointShapes } from '../core-shapes'
import {
	ILayout,
	IShadowOptions,
	ISlide,
	ISlideLayout,
	ISlideObject,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	ITableCell,
	ITableCellOpts,
	IObjectOptions,
	IText,
	ITextOpts,
} from '../core-interfaces'
import { encodeXmlEntities, inch2Emu, genXmlColorSelection, getSmartParseNumber, convertRotationDegrees } from '../gen-utils'

import Bullet from './bullet'
import Hyperlink from './hyperlink'

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

export default class TextFragment {
	text

	rtlMode

	bullet

	align
	lineSpacing
	indentLevel
	paraSpaceBefore
	paraSpaceAfter

	lang

	fontFace
	fontSize
	charSpacing
	color
	bold
	italic
	strike
	underline
	subscript
	superscript
	outline

	hyperlink

	constructor(text, options, registerLink) {
		this.text = text

		const { paraSpaceBefore, paraSpaceAfter, indentLevel } = options

		if (indentLevel && !isNaN(Number(indentLevel)) && Number(this.indentLevel) > 0) this.indentLevel = Number(indentLevel)
		if (paraSpaceBefore && !isNaN(Number(paraSpaceBefore)) && Number(this.paraSpaceBefore) > 0) this.paraSpaceBefore = Number(paraSpaceBefore)
		if (paraSpaceAfter && !isNaN(Number(paraSpaceAfter)) && Number(this.paraSpaceAfter) > 0) this.paraSpaceAfter = Number(paraSpaceAfter)

		this.bullet = new Bullet(options.bullet)

		if (options.hyperlink) {
			this.hyperlink = new Hyperlink(options.hyperlink, registerLink)
		}

		const alignInput = (options.align || '').toLowerCase()
		if (alignInput.startsWith('c')) this.align = TEXT_HALIGN.center
		else if (alignInput.startsWith('l')) this.align = TEXT_HALIGN.left
		else if (alignInput.startsWith('r')) this.align = TEXT_HALIGN.right
		else if (alignInput.startsWith('j')) this.align = TEXT_HALIGN.justify

		this.rtlMode = options.rtlMode

		this.lineSpacing = options.lineSpacing && !isNaN(options.lineSpacing) ? options.lineSpacing : null
		this.paraSpaceBefore = options.paraSpaceBefore
		this.paraSpaceAfter = options.paraSpaceAfter
		this.lang = options.lang
		this.fontFace = options.fontFace
		this.fontSize = options.fontSize
		this.charSpacing = options.charSpacing
		this.color = options.color
		this.bold = options.bold
		this.italic = options.italic
		this.strike = options.strike
		this.underline = options.underline
		this.subscript = options.subscript
		this.superscript = options.superscript
		this.outline = options.outline
	}

	render() {
		let bulletLvl0Margin = 342900

		// B: 'lstStyle'
		// NOTE: shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
		// FIXME: LINE horiz-align doesnt work (text is always to the left inside line) (FYI: the PPT code diff is substantial!)
		//if (opts.h === 0 && opts.line && opts.align) {
		//strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>'
		//} else if (slideObj.type === 'placeholder') {
		//strSlideXml += `<a:lstStyle>${genXmlParagraphProperties(slideObj, true)}</a:lstStyle>`
		//} else {
		//strSlideXml += '<a:lstStyle/>'
		//}

		const marginLeft = this.indentLevel && this.indentLevel > 0 ? bulletLvl0Margin + bulletLvl0Margin * this.indentLevel : bulletLvl0Margin

		return `
        <a:pPr ${[
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
        </a:pPr>
        <a:r>
            <a:rPr lang="${this.lang ? this.lang : 'en-US'}"${this.lang ? ' altLang="en-US"' : ''} ${[
			// NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
			this.fontSize ? ` sz="${Math.round(this.fontSize)}00"` : '',
			this.bold ? ' b="1"' : '',
			this.italic ? ' i="1"' : '',
			this.strike ? ' strike="sngStrike"' : '',
			this.underline || this.hyperlink ? ' u="sng"' : '',
			this.subscript ? ' baseline="-40000"' : this.superscript ? ' baseline="30000"' : '',
			// IMPORTANT: Also disable kerning; otherwise text won't actually expand
			this.charSpacing ? ` spc="${this.charSpacing * 100}" kern="0"` : '',
			'dirty="0"',
		].join('')} > ${[
			this.color ? genXmlColorSelection(this.color) : '',
			// NOTE: 'cs' = Complex Script, 'ea' = East Asian (use "-120" instead of "0" - per Issue #174); ea must come first (Issue #174)
			this.fontFace
				? `
                    <a:latin typeface="${this.fontFace}" pitchFamily="34" charset="0"/>
                    <a:ea typeface="${this.fontFace}" pitchFamily="34" charset="-122"/>
                    <a:cs typeface="${this.fontFace}" pitchFamily="34" charset="-120"/>`
				: '',
			this.outline ? `<a:ln w="${Math.round((this.outline.size || 0.75) * ONEPT)}">${genXmlColorSelection(this.outline.color || 'FFFFFF')}</a:ln>` : '',
			this.hyperlink ? this.hyperlink.render() : '',
		].join('')}
            </a:rPr>
            <a:t>${encodeXmlEntities(this.text)}</a:t>
        </a:r>
        `
	}
}
