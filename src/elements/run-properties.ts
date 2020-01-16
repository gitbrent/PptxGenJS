import { ONEPT } from '../core-enums'
import { genXmlColorSelection, translateColor } from '../gen-utils'
import Relations from '../relations'

import Hyperlink, { HyperLinkOptions } from './hyperlink'

export interface RunPropertiesOptions {
    lang?: string
    fontFace?: string
    fontSize?: number
    charSpacing?: number
    color?: string
    bold?: boolean
    italic?: boolean
    strike?: boolean
    underline?: boolean
    subscript?: boolean
    superscript?: boolean
    outline?: { size?: number; color?: string }
    hyperlink?: HyperLinkOptions
}

export default class RunProperties {
    lang: string
    altLang: string

    fontFace?: string
    fontSize?: number
    charSpacing?: number
    color?: string
    bold: boolean
    italic: boolean
    strike: boolean
    underline: boolean
    subscript: boolean
    superscript: boolean

    outline?: { size: number; color: string }
    hyperlink?: Hyperlink

    constructor(options: RunPropertiesOptions, relations: Relations) {
        this.lang = options.lang || 'en-US'
        this.altLang = options.lang ? '' : 'en-US'
        this.fontFace = options.fontFace
        // NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
        this.fontSize = options.fontSize && Math.round(options.fontSize)
        this.charSpacing = options.charSpacing
        this.color = translateColor(options.color)
        this.bold = !!options.bold
        this.italic = !!options.italic
        this.strike = !!options.strike
        this.underline = !!options.underline
        this.subscript = !!options.subscript
        this.superscript = !!options.superscript

        if (options.outline) {
            this.outline = {
                size: options.outline.size || 0.75,
                color: translateColor(options.outline.color) || 'FFFFFF'
            }
        }

        if (options.hyperlink) {
            this.hyperlink = new Hyperlink(options.hyperlink, relations)
        }
    }

    render(tag) {
        return `<${tag} ${[
            ` lang="${this.lang}"`,
            this.altLang ? ` altLang="${this.altLang}"` : '',
            this.fontSize ? ` sz="${this.fontSize}00"` : '',
            this.bold ? ' b="1"' : '',
            this.italic ? ' i="1"' : '',
            this.strike ? ' strike="sngStrike"' : '',
            this.underline || this.hyperlink ? ' u="sng"' : '',
            this.subscript
                ? ' baseline="-40000"'
                : this.superscript
                ? ' baseline="30000"'
                : '',
            // IMPORTANT: Also disable kerning; otherwise text won't actually expand
            this.charSpacing ? ` spc="${this.charSpacing * 100}" kern="0"` : '',
            ' dirty="0"'
        ].join('')} > ${[
            this.color ? genXmlColorSelection(this.color) : '',
            // NOTE: 'cs' = Complex Script, 'ea' = East Asian (use "-120" instead of "0" - per Issue #174); ea must come first (Issue #174)
            this.fontFace
                ? `
                    <a:latin typeface="${this.fontFace}" pitchFamily="34" charset="0"/>
                    <a:ea typeface="${this.fontFace}" pitchFamily="34" charset="-122"/>
                    <a:cs typeface="${this.fontFace}" pitchFamily="34" charset="-120"/>`
                : '',
            this.outline
                ? `<a:ln w="${Math.round(
                      this.outline.size * ONEPT
                  )}">${genXmlColorSelection(this.outline.color)}</a:ln>`
                : '',
            this.hyperlink ? this.hyperlink.render() : ''
        ].join('')}
      </${tag}>
`
    }
}
