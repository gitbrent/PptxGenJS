import { SLIDE_OBJECT_TYPES } from '../core-enums'
import Position from './position'

export default class Placeholder {
	type = SLIDE_OBJECT_TYPES.newtext

	public name: string
	public position: Position

	public placeholderType
	protected placeholderIndex

	constructor(name, type = 'body', index) {
		this.name = name
		this.placeholderType = type
		this.placeholderIndex = index
	}

	renderPlaceholderInfo() {
		return `<p:ph idx="${this.placeholderIndex}" type="${this.placeholderType}" />`
	}

	public render(idx, presLayout) {
		throw new Error('not implemented')
	}
}
