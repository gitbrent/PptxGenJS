import { ONEPT } from '../core-enums'
import { genXmlColorSelection, translateColor } from '../gen-utils'

const DASH_VALUES = [
    'dash',
    'dashDot',
    'lgDash',
    'lgDashDot',
    'lgDashDotDot',
    'solid',
    'sysDash',
    'sysDot'
]

// https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_LineCap_topic_ID0EPMVNB.html#topic_ID0EPMVNB
const CAP_VALUES = ['flat', 'rnd', 'sq']

export default class LineElement {
    size
    color
    dash
    head
    tail
    cap = 'sq'

    constructor({ size = 1, color = '333333', dash, head, tail, cap }) {
        this.color = translateColor(color)
        this.size = size

        if (!DASH_VALUES.includes(dash)) this.dash = 'solid'
        else this.dash = dash

        this.dash = dash
        this.head = head
        this.tail = tail

        // omitting cap if uknown value
        // assumed 'sq' when empty
        // https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ln_topic_ID0EJR3IB.html#topic_ID0EJR3IB
        if (CAP_VALUES.includes(cap)) this.cap = cap
    }

    render() {
        return `
    <a:ln${this.size ? ` w="${this.size * ONEPT}"` : ''} cap="${this.cap}">
        ${genXmlColorSelection(this.color)}
        ${this.dash ? `<a:prstDash val="${this.dash}"/>` : ''}
        ${this.head ? `<a:headEnd type="${this.head}"/>` : ''}
        ${this.tail ? `<a:tailEnd type="${this.tail}"/>` : ''}
    </a:ln>`
    }
}
