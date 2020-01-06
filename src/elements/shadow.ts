import { ONEPT } from '../core-enums'

export default class ShadowElement {
    type
    blur
    offset
    angle
    color
    opacity

    constructor(options) {
        const {
            type: inputType = 'outer',
            blur = 8,
            offset = 4,
            angle: inputAngle = 270,
            color = '000000',
            opacity: inputOpacity = 0.75
        } = options

        let type = inputType
        if (type !== 'outer' && type !== 'inner' && type !== 'none') {
            console.warn(
                'Warning: shadow.type options are `outer`, `inner` or `none`.'
            )
            type = 'outer'
        }

        let angle = inputAngle
        if (angle) {
            if (isNaN(Number(angle)) || angle < 0 || angle > 359) {
                console.warn('Warning: shadow.angle can only be 0-359')
                angle = 270
            }

            // B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
            angle = Math.round(Number(angle))
        }

        let opacity = inputOpacity
        if (opacity) {
            if (isNaN(Number(opacity)) || opacity < 0 || opacity > 1) {
                console.warn('Warning: shadow.opacity can only be 0-1')
                opacity = 0.75
            }

            // B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
            opacity = Number(opacity)
        }

        this.type = type
        this.blur = blur * ONEPT
        this.offset = offset * ONEPT
        this.angle = angle * 60000
        this.color = color
        this.opacity = opacity * 100000
    }

    render() {
        const tag = `a:${this.type}Shdw`
        return `
	<a:effectLst>
        <${tag} 
         sx="100000" 
         sy="100000" 
         kx="0" 
         ky="0" 
         algn="bl" 
         rotWithShape="0" 
         blurRad="${this.blur}" 
         dist="${this.offset}" 
         dir="${this.angle}">
			<a:srgbClr val="${this.color}">
			<a:alpha val="${this.opacity}"/></a:srgbClr>
		</${tag}>'
	</a:effectLst>`
    }
}
