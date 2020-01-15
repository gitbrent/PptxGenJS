import {
    CRLF,
    ONEPT,
    TEXT_VALIGN,
    DEF_FONT_COLOR,
    SLIDE_OBJECT_TYPES
} from '../core-enums'

import { inch2Emu, genXmlColorSelection } from '../gen-utils'
import Relations from '../relations'

import ElementInterface from './element-interface'

import ShadowElement, { ShadowOptions } from './shadow'
import Shape, { ShapeConfig, ShapeOptions } from './shape'
import Position, { PositionOptions } from './position'
import LineElement from './line'
import TextFragment, { buildFragments, FragmentOptions } from './text-fragment'

import ParagraphProperties, {
    ParagraphPropertiesOptions
} from './paragraph-properties'
import RunProperties, { RunPropertiesOptions } from './run-properties'

export type TextOptions = PositionOptions &
    ParagraphPropertiesOptions &
    RunPropertiesOptions &
    ShapeOptions & {
        breakLine?: boolean
        shape?: ShapeConfig
        fill?: string
        placeholder?: string
        autoFit?: boolean
        line?: string
        lineSize?: number
        lineDash?: string
        lineHead?: string
        lineTail?: string
        vert?: string
        isTextBox?: boolean
        shrinkText?: boolean
        inset?: number
        margin?: number | [number, number, number, number]
        valign?: string
        shadow?: ShadowOptions
        wrap?: boolean
    }

export default class TextElement implements ElementInterface {
    type = SLIDE_OBJECT_TYPES.newtext

    fragments: TextFragment[]
    shape: Shape
    fill
    color
    lang

    position: Position

    line
    lineSize

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
    placeholder?: string

    paragraphProperties: ParagraphProperties
    runProperties: RunProperties

    constructor(
        text: FragmentOptions,
        opts: TextOptions,
        relations: Relations
    ) {
        this.fragments = buildFragments(text, opts.breakLine, relations)
        if (!opts.placeholder || opts.shape) {
            this.shape = new Shape(opts.shape, { rectRadius: opts.rectRadius })
        } else {
            this.shape = null
        }

        this.fill = opts.fill
        this.lang = opts.lang

        if (opts.placeholder) this.placeholder = opts.placeholder

        // A: Placeholders should inherit their colors or override them, so don't default them
        if (!opts.placeholder) {
            this.color = opts.color || DEF_FONT_COLOR // Set color (options > inherit from Slide > default to black)
        }

        if (opts.line || (this.shape && this.shape.name === 'line')) {
            this.line = new LineElement({
                color: opts.line,
                size: opts.lineSize,
                dash: opts.lineDash,
                head: opts.lineHead,
                tail: opts.lineTail
            })
        }

        this.position = new Position({
            x: opts.x,
            y: opts.y,
            h: opts.h,
            w: opts.w,
            flipV: opts.flipV,
            flipH: opts.flipH,
            rotate: opts.rotate
        })

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
            paraSpaceAfter: opts.paraSpaceAfter
        })
        this.runProperties = new RunProperties(
            {
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
                hyperlink: opts.hyperlink
            },
            relations
        )
    }

    render(idx, presLayout, placeholder) {
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
            ${placeholder ? placeholder.renderPlaceholderInfo() : ''}
		    </p:nvPr>
        </p:nvSpPr>

        <p:spPr>
            ${this.position.render(presLayout)}
            ${this.shape ? this.shape.render(this.position, presLayout) : ''}
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
                this.vert ? `vert="${this.vert}"` : '' // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
            ].join(' ')}>
                ${
                    this.shrinkText
                        ? '<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>'
                        : ''
                }
                ${this.autoFit !== false ? '<a:spAutoFit/>' : ''}
            </a:bodyPr>

            <a:lstStyle>
                ${this.paragraphProperties.render(
                    presLayout,
                    'a:lvl1pPr',
                    this.runProperties.render('a:defRPr')
                )}
            </a:lstStyle>
            <a:p>
                ${this.fragments
                    .map(fragment => fragment.render(presLayout))
                    .join('</a:p><a:p>')}
                ${'' /* NOTE: Added 20180101 to address PPT-2007 issues */}
		        <a:endParaRPr lang="${this.lang || 'en-US'}" dirty="0"/>
            </a:p>
        </p:txBody>
    </p:sp>`
    }
}
