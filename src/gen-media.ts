/**
 * PptxGenJS: Media Methods
 */

import { IMG_BROKEN } from './core-enums'
import { ISlide, ISlideLayout, ISlideRelMedia } from './core-interfaces'

/**
 * Encode Image/Audio/Video into base64
 */
export function encodeSlideMediaRels(layout: ISlide | ISlideLayout): Promise<string>[] {
	const fs = typeof require !== 'undefined' ? require('fs') : null // NodeJS
	const https = typeof require !== 'undefined' ? require('https') : null // NodeJS
	let imageProms: Promise<string>[] = []

	// A: Read/Encode each audio/image/video thats not already encoded (eg: base64 provided by user)
	layout.relsMedia
		.filter(rel => {
			return rel.type != 'online' && !rel.data
		})
		.forEach(rel => {
			imageProms.push(
				new Promise((resolve, reject) => {
					if (fs && rel.path.indexOf('http') != 0) {
						// DESIGN: Node local-file encoding is syncronous, so we can load all images here, then call export with a callback (if any)
						try {
							let bitmap = fs.readFileSync(rel.path)
							rel.data = Buffer.from(bitmap).toString('base64')
							resolve('done')
						} catch (ex) {
							rel.data = IMG_BROKEN
							reject('ERROR: Unable to read media: "' + rel.path + '"\n' + ex.toString())
						}
					} else if (fs && https && rel.path.indexOf('http') == 0) {
						https.get(rel.path, res => {
							var rawData = ''
							res.setEncoding('binary') // IMPORTANT: Only binary encoding works
							res.on('data', chunk => (rawData += chunk))
							res.on('end', () => {
								rel.data = Buffer.from(rawData, 'binary').toString('base64')
								resolve('done')
							})
							res.on('error', ex => {
								rel.data = IMG_BROKEN
								reject('ERROR: Unable to load image: "' + rel.path + '"\n' + ex.toString())
							})
						})
					} else {
						// A: Declare XHR and onload/onerror handlers
						// DESIGN: `XMLHttpRequest()` plus `FileReader()` = Ablity to read any file into base64!
						let xhr = new XMLHttpRequest()
						xhr.onload = () => {
							let reader = new FileReader()
							reader.onloadend = () => {
								rel.data = reader.result
								if (!rel.isSvgPng) {
									resolve('done')
								} else {
									createSvgPngPreview(rel)
										.then(() => {
											resolve('done')
										})
										.catch(ex => {
											reject(ex.toString())
										})
								}
							}
							reader.readAsDataURL(xhr.response)
						}
						xhr.onerror = ex => {
							rel.data = IMG_BROKEN
							reject('ERROR: Unable to load image: "' + rel.path + '"\n' + ex.toString())
						}

						// B: Execute request
						xhr.open('GET', rel.path)
						xhr.responseType = 'blob'
						xhr.send()
					}
				})
			)
		})

	// B: SVG: base64 data still requires a png to be generated (`isSvgPng` flag this as the preview image, not the SVG itself)
	layout.relsMedia
		.filter(rel => {
			return rel.isSvgPng && rel.data
		})
		.forEach(rel => {
			if (fs) {
				console.log('Sorry, SVG is not supported in Node (more info: https://github.com/gitbrent/PptxGenJS/issues/401)')
				rel.data = IMG_BROKEN
				imageProms.push(Promise.resolve('done'))
			} else {
				imageProms.push(createSvgPngPreview(rel))
			}
		})

	return imageProms
}

function createSvgPngPreview(rel: ISlideRelMedia): Promise<string> {
	return new Promise((resolve, reject) => {
		// A: Create
		let image = new Image()

		// B: Set onload event
		image.onload = () => {
			// First: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (image.width + image.height == 0) {
				image.onerror('h/w=0')
			}
			let canvas: HTMLCanvasElement = document.createElement('CANVAS') as HTMLCanvasElement
			let ctx = canvas.getContext('2d')
			canvas.width = image.width
			canvas.height = image.height
			ctx.drawImage(image, 0, 0)
			// Users running on local machine will get the following error:
			// "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
			// when the canvas.toDataURL call executes below.
			try {
				rel.data = canvas.toDataURL(rel.type)
				resolve('done')
			} catch (ex) {
				image.onerror(ex)
			}
			canvas = null
		}
		image.onerror = ex => {
			rel.data = IMG_BROKEN
			reject(ex.toString())
		}

		// C: Load image
		image.src = typeof rel.data === 'string' ? rel.data : IMG_BROKEN
	})
}

/**
 * FIXME: TODO: currently unused
 * TODO: Should return a Promise
 */
function getSizeFromImage(inImgUrl: string): { width: number; height: number } {
	const sizeOf = typeof require !== 'undefined' ? require('sizeof') : null // NodeJS

	if (sizeOf) {
		try {
			let dimensions = sizeOf(inImgUrl)
			return { width: dimensions.width, height: dimensions.height }
		} catch (ex) {
			console.error('ERROR: sizeOf: Unable to load image: ' + inImgUrl)
			return { width: 0, height: 0 }
		}
	} else if (Image && typeof Image === 'function') {
		// A: Create
		let image = new Image()

		// B: Set onload event
		image.onload = () => {
			// FIRST: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (image.width + image.height == 0) {
				return { width: 0, height: 0 }
			}
			var obj = { width: image.width, height: image.height }
			return obj
		}
		image.onerror = () => {
			try {
				console.error('ERROR: image.onload: Unable to load image: ' + inImgUrl)
			} catch (ex) {}
		}

		// C: Load image
		image.src = inImgUrl
	}
}
