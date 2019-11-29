import { BULLET_TYPES } from '../core-enums'

export default class Bullet {
	enabled
	default
	inherit

	code
	type

	style
	startAt

	bulletCode

	constructor(bullet) {
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
		this.startAt = bullet.startAt || '1'

		// Check value for hex-ness (s/b 4 char hex)
		if (this.code && /^[0-9A-Fa-f]{4}$/.test(this.code) === false) {
			console.warn('Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!')
			this.bulletCode = BULLET_TYPES['DEFAULT']
		}
		this.bulletCode = this.code && `&#x${this.code};`

		this.enabled = this.code || this.type === 'number'

		this.style = bullet.style
	}

	render() {
		if (this.enabled && this.type === 'number') {
			return `<a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="${this.style}" startAt="${this.startAt}"/>`
		} else if (this.enabled && this.code) {
			return `<a:buSzPct val="100000"/><a:buChar char="${this.bulletCode}"/>`
		} else if (this.enabled && this.default) {
			return `<a:buSzPct val="100000"/><a:buChar char="${BULLET_TYPES['DEFAULT']}"/>`
		} else if (!this.enabled && !this.inherit) {
			return '<a:buNone/>'
		}

		return ''
	}
}
