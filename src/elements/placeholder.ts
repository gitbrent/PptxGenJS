import ElementInterface from './element-interface'
import Position from './position'

export type PlaceholderOptions = {
    name: string
    type: string
}

export default class Placeholder implements ElementInterface {
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
        return ''
    }
}
