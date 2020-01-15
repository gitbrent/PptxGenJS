import { genericParseFloat } from '../gen-utils'

import Placeholder, { PlaceholderOptions } from './placeholder'
import ImageElement, { ObjectFitOptions, ColorBlend } from './image'
import Position, { PositionOptions } from './position'

export type PlaceholderImageOptions = PositionOptions &
    PlaceholderOptions & {
        objectFit?: ObjectFitOptions
        colorBlend?: ColorBlend
        opacity?: number
    }

export default class PlaceholderImage extends Placeholder {
    public position: Position
    public objectFit?: ObjectFitOptions
    opacity?: number
    colorBlend?: ColorBlend

    constructor(options: PlaceholderImageOptions, index) {
        super(options.name, options.type || 'pic', index)

        this.position = new Position({
            x: options.x,
            y: options.y,
            h: options.h,
            w: options.w,

            flipV: options.flipV,
            flipH: options.flipH,
            rotate: options.rotate
        })

        this.objectFit = options.objectFit
        this.colorBlend = options.colorBlend

        if (options.opacity) {
            const numberOpacity = genericParseFloat(options.opacity)
            if (numberOpacity < 1 && numberOpacity >= 0) {
                this.opacity = numberOpacity
            }
        }
    }

    render(idx, presLayout) {
        return `
    <p:sp>
        <p:nvSpPr>
            <p:cNvPr id="${idx + 2}" name="Placeholder ${idx + 1}"/>
            <p:cNvSpPr />
		    <p:nvPr>
            ${this.renderPlaceholderInfo()}
		    </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            ${this.position.render(presLayout)}
		    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
    </p:sp>
    `
    }
}
