import { ONEPT, TEXT_VALIGN } from '../core-enums'
import { inch2Emu } from '../gen-utils'

export type VerticalOptions =
    | 'eaVert'
    | 'horz'
    | 'mongolianVert'
    | 'vert'
    | 'vert270'
    | 'wordArtVert'
    | 'wordArtVertRtl'

export interface BodyPropertiesOptions {
    autoFit?: boolean
    vert?: VerticalOptions
    shrinkText?: boolean
    inset?: number
    margin?: number | [number, number, number, number]
    valign?: string
    wrap?: string
}

export default class BodyProperties {
    autoFit: boolean
    shrinkText: boolean
    anchor: TEXT_VALIGN
    vert?: VerticalOptions

    lIns: number
    rIns: number
    tIns: number
    bIns: number

    wrap?: string

    constructor(opts: BodyPropertiesOptions, disableDefaults: boolean) {
        // D: Transform text options to bodyProperties as thats how we build XML
        this.autoFit = opts.autoFit || false // If true, shape will collapse to text size (Fit To shape)
        this.shrinkText = opts.shrinkText || false
        this.anchor = disableDefaults ? null : TEXT_VALIGN.ctr
        this.vert = opts.vert || null

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

        this.wrap = opts.wrap || (!disableDefaults && 'square')
    }

    render() {
        // NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
        // MS-PPT > Format shape > Text Options: "Shrink text on overflow"
        // MS-PPT > Format shape > Text Options: "Resize shape to fit text" [spAutoFit]
        return `<a:bodyPr ${[
            this.wrap ? `wrap="${this.wrap}"` : '',
            this.lIns || this.lIns === 0 ? `lIns="${this.lIns}"` : '',
            this.tIns || this.tIns === 0 ? `tIns="${this.tIns}"` : '',
            this.rIns || this.rIns === 0 ? `rIns="${this.rIns}"` : '',
            this.bIns || this.bIns === 0 ? `bIns="${this.bIns}"` : '',
            'rtlCol="0"',
            this.anchor ? `anchor="${this.anchor}"` : '',
            this.vert ? `vert="${this.vert}"` : ''
        ].join(' ')}>${
            this.shrinkText
                ? '<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>'
                : ''
        }${this.autoFit ? '<a:spAutoFit/>' : ''}
          </a:bodyPr>`
    }
}
