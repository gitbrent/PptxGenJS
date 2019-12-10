import { CRLF, ONEPT, TEXT_VALIGN, DEF_FONT_COLOR, SLIDE_OBJECT_TYPES } from '../core-enums'

import { inch2Emu, genXmlColorSelection } from '../gen-utils'

import ShadowElement from './shadow'
import Shape from './shape'
import Position from './position'
import LineElement from './line'
import TextFragment from './text-fragment'

import ParagraphProperties from './paragraph-properties'
import RunProperties from './run-properties'
import Hyperlink from './hyperlink'

const buildFragments = (inputText, inputBreakLine, registerLink) => {
	let fragments = inputText
	if (typeof inputText === 'string' || typeof inputText === 'number') {
		fragments = [{ text: inputText.toString(), options: {} }]
	}
	if (!fragments) return []

	return fragments.flatMap(({ text: fragmentText, options }, idx) => {
		let text = fragmentText.replace(/\r*\n/g, CRLF)
		let breakLine = inputBreakLine || options.breakLine || false

		if (text.indexOf(CRLF) > -1) {
			// Remove trailing linebreak (if any) so the "if" below doesnt create a double CRLF+CRLF line ending!
			text = text.replace(/\r\n$/g, '')
			// Plain strings like "hello \n world" or "first line\n" need to have lineBreaks set to become 2 separate lines as intended
			breakLine = true
		}

		const paragraphProperties = new ParagraphProperties({
			bullet: options.bullet,
			align: options.align,
			rtlMode: options.rtlMode,
			lineSpacing: options.lineSpacing,
			indentLevel: options.indentLevel,
			paraSpaceBefore: options.paraSpaceBefore,
			paraSpaceAfter: options.paraSpaceAfter,
		})
		const runProperties = new RunProperties({
			lang: options.lang,
			fontFace: options.fontFace,
			fontSize: options.fontSize,
			charSpacing: options.charSpacing,
			color: options.color,
			bold: options.bold,
			italic: options.italic,
			strike: options.strike,
			underline: options.underline,
			subscript: options.subscript,
			superscript: options.superscript,
			outline: options.outline,
			hyperlink: options.hyperlink && new Hyperlink(options.hyperlink, registerLink),
		})

		if (breakLine) {
			return text.split(CRLF).map((line, lineIdx) => {
				return new TextFragment(line, paragraphProperties, runProperties)
			})
		} else {
			return new TextFragment(text, paragraphProperties, runProperties)
		}
	})
}

export default class TextElement {
	type = SLIDE_OBJECT_TYPES.newtext

	fragments
	shape
	fill
	color
	lang

	position

	line
	lineSize

	rectRadius

	autoFit
	shrinkText
	anchor
	vert

	isTextBox

	lIns
	rIns
	tIns
	bIns

	valign
	wrap

	shadow
	placeholder

	paragraphProperties
	runProperties

	constructor(text, opts, registerLink) {
		this.fragments = buildFragments(text, opts.breakLine, registerLink)
		if (!opts.placeholder || opts.shape) {
			this.shape = new Shape(opts.shape)
		} else {
			this.shape = null
		}

		this.fill = opts.fill
		this.lang = opts.lang

		this.placeholder = opts.placeholder

		// A: Placeholders should inherit their colors or override them, so don't default them
		if (!opts.placeholder) {
			this.color = opts.color || DEF_FONT_COLOR // Set color (options > inherit from Slide > default to black)
		}

		if (opts.line || (opts.shape && opts.shape.name === 'line')) {
			this.line = new LineElement({
				color: opts.line,
				size: opts.lineSize,
				dash: opts.lineDash,
				head: opts.lineHead,
				tail: opts.lineTail,
			})
		}

		this.position = new Position({
			x: opts.x,
			y: opts.y,
			h: opts.h,
			w: opts.w,
			flipV: opts.flipV,
			flipH: opts.flipH,
			rotate: opts.rotate,
		})

		this.rectRadius = opts.rectRadius

		// D: Transform text options to bodyProperties as thats how we build XML
		this.autoFit = opts.autoFit || false // If true, shape will collapse to text size (Fit To shape)
		this.shrinkText = opts.shrinkText || false
		this.anchor = opts.placeholder ? null : TEXT_VALIGN.ctr // VALS: [t,ctr,b]
		this.vert = opts.vert || null // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]

		this.isTextBox = opts.isTextBox

		// Margin/Padding/Inset for textboxes
		if ((opts.inset && !isNaN(Number(opts.inset))) || opts.inset === 0) {
			const inset = inch2Emu(opts.inset)
			this.lIns = inset
			this.rIns = inset
			this.tIns = inset
			this.bIns = inset
		}
		if (opts.margin && Array.isArray(opts.margin)) {
			this.lIns = opts.margin[0] * ONEPT || 0
			this.rIns = opts.margin[1] * ONEPT || 0
			this.bIns = opts.margin[2] * ONEPT || 0
			this.tIns = opts.margin[3] * ONEPT || 0
		} else if (typeof opts.margin === 'number') {
			const marginSize = opts.margin * ONEPT
			this.lIns = marginSize
			this.rIns = marginSize
			this.bIns = marginSize
			this.tIns = marginSize
		}

		const valignInput = (opts.valign || '').toLowerCase()
		if (valignInput.startsWith('b')) this.anchor = TEXT_VALIGN.b
		else if (valignInput.startsWith('c')) this.anchor = TEXT_VALIGN.ctr
		else if (valignInput.startsWith('m')) this.anchor = TEXT_VALIGN.ctr
		else if (valignInput.startsWith('t')) this.anchor = TEXT_VALIGN.t

		this.wrap = opts.wrap || (opts.placeholder && 'square')

		if (opts.shadow) {
			this.shadow = new ShadowElement(opts.shadow)
		}

		this.paragraphProperties = new ParagraphProperties({
			bullet: opts.bullet,
			align: opts.align,
			rtlMode: opts.rtlMode,
			lineSpacing: opts.lineSpacing,
			indentLevel: opts.indentLevel,
			paraSpaceBefore: opts.paraSpaceBefore,
			paraSpaceAfter: opts.paraSpaceAfter,
		})
		this.runProperties = new RunProperties({
			lang: opts.lang,
			fontFace: opts.fontFace,
			fontSize: opts.fontSize,
			charSpacing: opts.charSpacing,
			color: opts.color,
			bold: opts.bold,
			italic: opts.italic,
			strike: opts.strike,
			underline: opts.underline,
			subscript: opts.subscript,
			superscript: opts.superscript,
			outline: opts.outline,
			hyperlink: opts.hyperlink && new Hyperlink(opts.hyperlink, registerLink),
		})
	}

	render(idx, presLayout, renderPlaceholder) {
		// F: NEW: Add autofit type tags
		// MS-PPT > Format shape > Text Options: "Shrink text on overflow"

		// MS-PPT > Format shape > Text Options: "Resize shape to fit text" [spAutoFit]
		// NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
		return `
    <p:sp>
        <p:nvSpPr>
            <p:cNvPr id="${idx + 2}" name="Object ${idx + 1}"/>
            <p:cNvSpPr${this.isTextBox ? ' txBox="1"' : ''}/>
		    <p:nvPr>
                ${renderPlaceholder(this.placeholder)}
		    </p:nvPr>
        </p:nvSpPr>

        <p:spPr>
            ${this.position.render(presLayout)}
            ${this.shape ? this.shape.render(this.rectRadius, this.position, presLayout) : ''}
            ${
				this.fill
					? genXmlColorSelection(this.fill)
					: // We only default to no fill if we have not specified a placeholder
					this.placeholder
					? ''
					: '<a:noFill/>'
			}
            ${this.line ? this.line.render() : ''}
            ${this.shadow ? this.shadow.render() : ''}
		</p:spPr>
        <p:txBody>
            <a:bodyPr ${[
				this.wrap ? `wrap="${this.wrap}"` : '',
				this.lIns || this.lIns === 0 ? `lIns="${this.lIns}"` : '',
				this.tIns || this.tIns === 0 ? `tIns="${this.tIns}"` : '',
				this.rIns || this.rIns === 0 ? `rIns="${this.rIns}"` : '',
				this.bIns || this.bIns === 0 ? `bIns="${this.bIns}"` : '',
				'rtlCol="0"',
				this.anchor ? `anchor="${this.anchor}"` : '', // VALS: [t,ctr,b]
				this.vert ? `vert="${this.vert}"` : '', // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
			].join(' ')}>
                ${this.shrinkText ? '<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>' : ''}
                ${this.autoFit !== false ? '<a:spAutoFit/>' : ''}
            </a:bodyPr>

            <a:lstStyle>
                ${this.paragraphProperties.render('a:lvl1pPr', this.runProperties.render('a:defRPr'))}
            </a:lstStyle>
            <a:p>
                ${this.fragments.map(fragment => fragment.render()).join('</a:p><a:p>')}
                ${'' /* NOTE: Added 20180101 to address PPT-2007 issues */}
		        <a:endParaRPr lang="${this.lang || 'en-US'}" dirty="0"/>
            </a:p>
        </p:txBody>
    </p:sp>`
	}
}
