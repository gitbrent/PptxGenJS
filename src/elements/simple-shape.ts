import { CRLF, ONEPT, TEXT_VALIGN, DEF_FONT_COLOR } from '../core-enums'

import { inch2Emu, genXmlColorSelection } from '../gen-utils'

import ElementInterface from './element-interface'

import ShadowElement from './shadow'
import Shape from './shape'
import Position from './position'
import Line from './line'

const defaultsToOne = x => x || (x === 0 ? 0 : 1)

export default class SimpleShapeElement implements ElementInterface {
    shape
    fill
    color

    position

    line
    lineSize

    rectRadius

    shadow

    constructor(shape, opts) {
        this.shape = new Shape(shape)

        this.fill = opts.fill
        this.rectRadius = opts.rectRadius

        if (opts.line || shape.name === 'line') {
            this.line = new Line({
                color: opts.line || '333333',
                size: opts.lineSize,
                dash: opts.lineDash,
                head: opts.lineHead,
                tail: opts.lineTail
            })
        }

        this.position = new Position({
            x: defaultsToOne(opts.x),
            y: defaultsToOne(opts.y),
            h: defaultsToOne(opts.h),
            w: defaultsToOne(opts.w),
            flipV: opts.flipV,
            flipH: opts.flipH,
            rotate: opts.rotate
        })

        if (opts.shadow) {
            this.shadow = new ShadowElement(opts.shadow)
        }
    }

    render(idx, presLayout) {
        return `
    <p:sp>
        <p:nvSpPr>
            <p:cNvPr id="${idx + 2}" name="Object ${idx + 1}"/>
            <p:cNvSpPr />
		    <p:nvPr />
        </p:nvSpPr>

        <p:spPr>
            ${this.position.render(presLayout)}
            ${this.shape.render(this.rectRadius, this.position, presLayout)}
            ${this.fill ? genXmlColorSelection(this.fill) : '<a:noFill/>'}
            ${this.line ? this.line.render() : ''}
            ${this.shadow ? this.shadow.render() : ''}
		</p:spPr>
    </p:sp>`
    }
}
