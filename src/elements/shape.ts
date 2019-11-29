import { PowerPointShapes } from '../core-shapes'
import { EMU, PLACEHOLDER_TYPES } from '../core-enums'

export function genXmlPlaceholder(placeholderObj) {
	if (!placeholderObj) return ''

	let placeholderIdx = placeholderObj.options && placeholderObj.options.placeholderIdx ? placeholderObj.options.placeholderIdx : ''
	let placeholderType = placeholderObj.options && placeholderObj.options.placeholderType ? placeholderObj.options.placeholderType : ''

	return `<p:ph
		${placeholderIdx ? ' idx="' + placeholderIdx + '"' : ''}
		${placeholderType && PLACEHOLDER_TYPES[placeholderType] ? ' type="' + PLACEHOLDER_TYPES[placeholderType] + '"' : ''}
		${placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : ''}
		/>`
}

export default class ShapeElement {
	displayName
	name
	avLst

	constructor(input) {
		let shapeConfig = input
		if (!input) shapeConfig = PowerPointShapes.RECTANGLE

		if (typeof input === 'string') {
			if (PowerPointShapes[input]) {
				shapeConfig = PowerPointShapes[input]
			}

			shapeConfig = Object.keys(PowerPointShapes).filter(key => {
				return PowerPointShapes[key].name === input || PowerPointShapes[key].displayName
			})[0]
		}

		if (!shapeConfig) shapeConfig = PowerPointShapes.RECTANGLE

		this.displayName = shapeConfig.displayName
		this.name = shapeConfig.name
		this.avLst = shapeConfig.avLst
	}

	render(rectRadius, position, presLayout) {
		const radius = rectRadius && Math.round((rectRadius * EMU * 100000) / Math.min(position.cx(presLayout), position.cy(presLayout)))
		return `
            <a:prstGeom prst="${this.name}">
                <a:avLst>${rectRadius ? `<a:gd name="adj" fmla="val ${radius}"/>` : ''}</a:avLst>
            </a:prstGeom>
        `
	}
}
