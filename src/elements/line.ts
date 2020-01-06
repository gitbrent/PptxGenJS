import { ONEPT } from '../core-enums'
import { genXmlColorSelection } from '../gen-utils'

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

export default class LineElement {
    size
    color
    dash
    head
    tail

    constructor({ size = 1, color = '333333', dash, head, tail }) {
        this.color = color
        this.size = size

        if (!DASH_VALUES.includes(dash)) this.dash = 'solid'
        else this.dash = dash

        this.dash = dash
        this.head = head
        this.tail = tail
    }

    render() {
        return `
    <a:ln${this.size ? ` w="${this.size * ONEPT}"` : ''}>
        ${genXmlColorSelection(this.color)}
        ${this.dash ? `<a:prstDash val="${this.dash}"/>` : ''}
        ${this.head ? `<a:headEnd type="${this.head}"/>` : ''}
        ${this.tail ? `<a:tailEnd type="${this.tail}"/>` : ''}
	</a:ln>`
    }
}
