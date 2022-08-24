import { FontProps, FontStyleProps, PresFont, PresFontStyle } from './core-interfaces'

export const encodePresFontRels = (fonts: PresFont[]): Promise<string>[] =>
	fonts
		.map(f => f.styles)
		.flat()
		.filter(v => !v.data && v.path)
		.map(loadFontStyle)

const loadFontStyle = (style: PresFontStyle): Promise<string> =>
	new Promise((resolve, reject) => {
		const xhr = new XMLHttpRequest()

		xhr.onload = () => {
			const reader = new FileReader()
			reader.onloadend = () => {
				style.data = reader.result
				resolve('done')
			}
			reader.readAsDataURL(xhr.response)
		}

		xhr.onerror = ex => {
			reject(`ERROR! Unable to load font (xhr.onerror): ${style.path}`)
		}

		xhr.open('GET', style.path)
		xhr.responseType = 'blob'
		xhr.send()
	})

const validStyles = ['regular', 'bold', 'italic', 'boldItalic']

const isValidStyle = (style: FontStyleProps) => typeof style === 'object' && validStyles.includes(style.name) && (style.path || style.data)

const isValidStyles = (styles: FontStyleProps[]) => Array.isArray(styles) && styles.every(isValidStyle)

const isValidFont = (font: FontProps) => typeof font === 'object' && font.name !== '' && isValidStyles(font.styles)

export const isValidFonts = (fonts: FontProps[]) => Array.isArray(fonts) && fonts.every(isValidFont)
