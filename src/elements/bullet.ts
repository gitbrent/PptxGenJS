import { BULLET_TYPES } from '../core-enums'
import { getSmartParseNumber, translateColor } from '../gen-utils'

export type BulletOptions =
    | boolean
    | {
          code?: string
          type?: string
          style?: string
          startAt?: number
          color?: string
          indent?: string | number
      }

export default class Bullet {
    enabled: boolean
    default: boolean
    inherit: boolean

    code?: string
    type?: string

    style?: string
    startAt?: number

    bulletCode?: string
    color?: string
    indent?: string | number

    constructor(bullet?: BulletOptions) {
        if (bullet === false) {
            this.enabled = false
            return
        }

        if (!bullet && bullet !== false) {
            this.inherit = true
            this.enabled = false
            return
        }

        if (bullet === true) {
            this.enabled = true
            this.default = true
            return
        }

        this.code = bullet.code
        this.type = bullet.type.toString().toLowerCase()

        this.style = bullet.style || 'arabicPeriod'
        this.startAt = bullet.startAt || 1

        // Check value for hex-ness (s/b 4 char hex)
        if (this.code && /^[0-9A-Fa-f]{4}$/.test(this.code) === false) {
            console.warn(
                'Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!'
            )
            this.bulletCode = BULLET_TYPES['DEFAULT']
        }
        this.bulletCode = this.code && `&#x${this.code};`
        this.color = translateColor(bullet.color)

        this.enabled = !!this.code || this.type === 'number'
        this.indent = bullet.indent
    }

    renderIndentProps(presLayout, indentLevel) {
        if (!this.enabled) return ''

        let bulletLvl0Margin = this.indent
            ? getSmartParseNumber(this.indent, 'X', presLayout)
            : 342900
        const marginLeft =
            indentLevel && indentLevel > 0
                ? bulletLvl0Margin + bulletLvl0Margin * indentLevel
                : bulletLvl0Margin

        return ` marL="${marginLeft}" indent="-${bulletLvl0Margin}"`
    }

    render() {
        const color = this.color
            ? `<a:buClr><a:srgbClr val="${this.color}"/></a:buClr>`
            : ''

        if (this.enabled && this.type === 'number') {
            return `${color}<a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="${this.style}" startAt="${this.startAt}"/>`
        } else if (this.enabled && this.code) {
            return `${color}<a:buSzPct val="100000"/><a:buChar char="${this.bulletCode}"/>`
        } else if (this.enabled && this.default) {
            return `${color}<a:buSzPct val="100000"/><a:buChar char="${BULLET_TYPES['DEFAULT']}"/>`
        } else if (!this.enabled && !this.inherit) {
            return '<a:buNone/>'
        }

        return ''
    }
}
