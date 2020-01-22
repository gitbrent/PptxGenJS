import { CRLF, ONEPT, TEXT_VALIGN, DEF_FONT_COLOR } from '../core-enums'

import { inch2Emu, translateColor, genXmlColorSelection } from '../gen-utils'
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
import BodyProperties, { BodyPropertiesOptions } from './body-properties'

export type TextOptions = PositionOptions &
    ParagraphPropertiesOptions &
    RunPropertiesOptions &
    BodyPropertiesOptions &
    ShapeOptions & {
        breakLine?: boolean
        shape?: ShapeConfig
        fill?: string
        placeholder?: string
        line?: string
        lineSize?: number
        lineDash?: string
        lineHead?: string
        lineTail?: string
        lineCap?: string
        isTextBox?: boolean
        shadow?: ShadowOptions
    }

export default class TextElement implements ElementInterface {
    fragments: TextFragment[]
    shape: Shape
    fill?: string
    color?: string
    lang?: string

    position: Position

    line: LineElement

    isTextBox?: boolean

    shadow
    placeholder?: string

    bodyProperties: BodyProperties
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

        this.fill = translateColor(opts.fill)
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
                tail: opts.lineTail,
                cap: opts.lineCap
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

        if (opts.shadow) {
            this.shadow = new ShadowElement(opts.shadow)
        }

        this.bodyProperties = new BodyProperties(
            {
                autoFit: opts.autoFit,
                vert: opts.vert,
                shrinkText: opts.shrinkText,
                inset: opts.inset,
                margin: opts.margin,
                valign: opts.valign,
                wrap: opts.wrap
            },
            !opts.placeholder
        )
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

    private renderFill() {
        if (this.fill) {
            return genXmlColorSelection(this.fill)
        }
        // We only default to no fill if we have not specified a placeholder
        if (this.placeholder) return ''
        return '<a:noFill/>'
    }

    render(idx, presLayout, placeholder) {
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
            ${this.renderFill()}
            ${this.line ? this.line.render() : ''}
            ${this.shadow ? this.shadow.render() : ''}
		    </p:spPr>
        <p:txBody>
            ${this.bodyProperties.render()}

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
