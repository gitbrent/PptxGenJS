import { PowerPointShapes } from '../core-shapes'
import { EMU } from '../core-enums'
import Position from './position'

export type ShapeConfig =
    | string
    | {
          displayName: string
          name: string
          avLst: { [key: string]: number }
      }

export type ShapeOptions = {
    rectRadius?: number
}

export default class ShapeElement {
    displayName: string
    name: string
    avLst: string

    rectRadius?: number

    constructor(input?: ShapeConfig, options: ShapeOptions = {}) {
        let shapeConfig
        if (!input) shapeConfig = PowerPointShapes.RECTANGLE

        if (typeof input === 'string') {
            if (PowerPointShapes[input]) {
                shapeConfig = PowerPointShapes[input]
            }

            shapeConfig = Object.keys(PowerPointShapes).filter(key => {
                return (
                    PowerPointShapes[key].name === input ||
                    PowerPointShapes[key].displayName
                )
            })[0]
        } else {
            shapeConfig = input
        }

        if (!shapeConfig) shapeConfig = PowerPointShapes.RECTANGLE

        this.displayName = shapeConfig.displayName
        this.name = shapeConfig.name
        this.avLst = shapeConfig.avLst

        if (options.rectRadius) {
            this.rectRadius = options.rectRadius * EMU * 100000
        }
    }

    private renderRadius(position, presLayout) {
        if (!this.rectRadius) return ''

        const smallerSide = Math.min(
            position.cx(presLayout),
            position.cy(presLayout)
        )
        const radius =
            this.rectRadius && Math.round(this.rectRadius / smallerSide)

        return `<a:gd name="adj" fmla="val ${radius}"/>`
    }

    render(position: Position, presLayout) {
        return `
            <a:prstGeom prst="${this.name}">
                <a:avLst>${this.renderRadius(position, presLayout)}</a:avLst>
            </a:prstGeom>
        `
    }
}
