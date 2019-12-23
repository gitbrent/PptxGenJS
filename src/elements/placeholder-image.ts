import { SLIDE_OBJECT_TYPES } from '../core-enums'
import Placeholder from './placeholder'
import ImageElement from './image'
import Position from './position'

export default class PlaceholderImage extends Placeholder {
	type = SLIDE_OBJECT_TYPES.newtext

	public position: Position
	public objectFit?: 'cover' | 'contain' | 'fill' | 'none'
	opacity?: number
	colorBlend?

	constructor(options, index) {
		super(options.name, options.type || 'pic', index)

		this.position = new Position({
			x: options.x,
			y: options.y,
			h: options.h,
			w: options.w,

			flipV: options.flipV,
			flipH: options.flipH,
			rotate: options.rotate,
		})

		this.objectFit = options.objectFit
		this.colorBlend = options.colorBlend

		if (options.opacity) {
			const numberOpacity = parseFloat(options.opacity)
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
