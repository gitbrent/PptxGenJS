import { SLIDE_OBJECT_TYPES } from '../core-enums'
import TextElement from './text'

export default class PlaceholderText {
	type = SLIDE_OBJECT_TYPES.newtext

	name
	textElement
	placeholderType
	placeholderIndex

	constructor(text, options, index, registerLink) {
		const { name, type = 'body', ...textOptions } = options

		// We default to no bullet in the placeholder (different from the slide
		// that inherits by default)
		if (!textOptions.bullet) textOptions.bullet = false

		const textElement = new TextElement(text, textOptions, registerLink)

		this.name = name
		this.textElement = textElement
		this.placeholderType = type
		this.placeholderIndex = index
	}

	renderPlaceholderInfo() {
		return `<p:ph idx="${this.placeholderIndex}" type="${this.placeholderType}" ${this.textElement.fragments.length > 0 ? ' hasCustomPrompt="1"' : ''} />`
	}

	render(idx, presLayout) {
		return this.textElement.render(idx, presLayout, this.renderPlaceholderInfo.bind(this))
	}
}
