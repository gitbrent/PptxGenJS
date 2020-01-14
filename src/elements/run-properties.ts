import { ONEPT } from '../core-enums'
import { genXmlColorSelection } from '../gen-utils'

export default class RunProperties {
    lang
    altLang

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

    constructor(options) {
        this.lang = options.lang || 'en-US'
        this.altLang = options.lang ? '' : 'en-US'
        this.fontFace = options.fontFace
        // NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
        this.fontSize = options.fontSize && Math.round(options.fontSize)
        this.charSpacing = options.charSpacing
        this.color = options.color
        this.bold = options.bold
        this.italic = options.italic
        this.strike = options.strike
        this.underline = options.underline
        this.subscript = options.subscript
        this.superscript = options.superscript
        this.outline = options.outline

        if (options.hyperlink) {
            this.hyperlink = options.hyperlink
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
                      (this.outline.size || 0.75) * ONEPT
                  )}">${genXmlColorSelection(
                      this.outline.color || 'FFFFFF'
                  )}</a:ln>`
                : '',
            this.hyperlink ? this.hyperlink.render() : ''
        ].join('')}
      </${tag}>
`
    }
}