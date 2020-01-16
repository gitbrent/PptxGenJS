import { CRLF, ONEPT, TEXT_VALIGN, DEF_FONT_COLOR } from '../core-enums'

import { inch2Emu, genXmlColorSelection, translateColor } from '../gen-utils'

import ElementInterface from './element-interface'

import ShadowElement, { ShadowOptions } from './shadow'
import Shape, { ShapeOptions as SO, ShapeConfig } from './shape'
import Position, { PositionOptions } from './position'
import Line from './line'

const defaultsToOne = x => x || (x === 0 ? 0 : 1)

type FullColor = string | { type: string; color: string; alpha?: number }

export type ShapeOptions = PositionOptions &
    SO & {
        fill?: FullColor
        color?: string
        rectRadius?: number
        line?: string
        lineSize?: number
        lineDash?: string
        lineHead?: string
        lineTail?: string
        shadow?: ShadowOptions
    }

export default class SimpleShapeElement implements ElementInterface {
    shape: Shape
    fill?: FullColor
    color?: string

    position: Position
    line?: Line

    shadow?: ShadowElement

    constructor(shape: ShapeConfig, opts: ShapeOptions) {
        this.shape = new Shape(shape, { rectRadius: opts.rectRadius })

        this.fill = translateColor(opts.fill)

        if (opts.line || this.shape.name === 'line') {
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
            ${this.shape.render(this.position, presLayout)}
            ${this.fill ? genXmlColorSelection(this.fill) : '<a:noFill/>'}
            ${this.line ? this.line.render() : ''}
            ${this.shadow ? this.shadow.render() : ''}
		</p:spPr>
    </p:sp>`
    }
}
