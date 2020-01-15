import Placeholder from './placeholder'
import Position from './Position'

export default interface Element {
    placeholder?: string
    render: (
        index: number,
        presLayout: any,
        placeholder?: Placeholder = null
    ) => string
}
