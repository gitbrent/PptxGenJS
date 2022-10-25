/**
 * PptxGenJS: Media Methods
 */

import { IMG_BROKEN } from './core-enums'
import { PresSlide, SlideLayout, ISlideRelMedia } from './core-interfaces'

/**
 * Encode Image/Audio/Video into base64
 * @param {PresSlide | SlideLayout} layout - slide layout
 * @return {Promise} promise
 */
export function encodeSlideMediaRels (layout: PresSlide | SlideLayout): Array<Promise<string>> {
	const fs = typeof require !== 'undefined' && typeof window === 'undefined' ? require('fs') : null // NodeJS
	const https = typeof require !== 'undefined' && typeof window === 'undefined' ? require('https') : null // NodeJS
	const imageProms: Array<Promise<string>> = []

	// A: Capture all audio/image/video candidates for encoding (filtering online/pre-encoded)
	const candidateRels = layout._relsMedia.filter(rel => rel.type !== 'online' && !rel.data && (!rel.path || (rel.path && !rel.path.includes('preencoded'))))

	// B: PERF: Mark dupes (same `path`) so that we dont load same media over-and-over
	const unqPaths: string[] = []
	candidateRels.forEach(rel => {
		if (!unqPaths.includes(rel.path)) {
			rel.isDuplicate = false
			unqPaths.push(rel.path)
		} else {
			rel.isDuplicate = true
		}
	})

	// C: Read/Encode each unique audio/image/video path
	candidateRels
		.filter(rel => !rel.isDuplicate)
		.forEach(rel => {
			imageProms.push(
				new Promise((resolve, reject) => {
					if (fs && rel.path.indexOf('http') !== 0) {
						// DESIGN: Node local-file encoding is syncronous, so we can load all images here, then call export with a callback (if any)
						try {
							const bitmap = fs.readFileSync(rel.path)
							rel.data = Buffer.from(bitmap).toString('base64')
							candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
							resolve('done')
						} catch (ex) {
							rel.data = IMG_BROKEN
							candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
							reject(new Error(`ERROR: Unable to read media: "${rel.path}"\n${String(ex)}`))
						}
					} else if (fs && https && rel.path.indexOf('http') === 0) {
						https.get(rel.path, (res) => {
							let rawData = ''
							res.setEncoding('binary') // IMPORTANT: Only binary encoding works
							res.on('data', (chunk: string) => (rawData += chunk))
							res.on('end', () => {
								rel.data = Buffer.from(rawData, 'binary').toString('base64')
								candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
								resolve('done')
							})
							res.on('error', (_ex) => {
								rel.data = IMG_BROKEN
								candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
								reject(new Error(`ERROR! Unable to load image (https.get): ${rel.path}`))
							})
						})
					} else {
						// A: Declare XHR and onload/onerror handlers
						// DESIGN: `XMLHttpRequest()` plus `FileReader()` = Ablity to read any file into base64!
						const xhr = new XMLHttpRequest()
						xhr.onload = () => {
							const reader = new FileReader()
							reader.onloadend = () => {
								rel.data = reader.result
								candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
								if (!rel.isSvgPng) {
									resolve('done')
								} else {
									createSvgPngPreview(rel)
										.then(() => {
											resolve('done')
										})
										.catch(ex => {
											reject(ex)
										})
								}
							}
							reader.readAsDataURL(xhr.response)
						}
						xhr.onerror = ex => {
							rel.data = IMG_BROKEN
							candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
							reject(new Error(`ERROR! Unable to load image (xhr.onerror): ${rel.path}`))
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
	layout._relsMedia
		.filter(rel => rel.isSvgPng && rel.data)
		.forEach(rel => {
			if (fs) {
				// console.log('Sorry, SVG is not supported in Node (more info: https://github.com/gitbrent/PptxGenJS/issues/401)')
				rel.data = IMG_BROKEN
				imageProms.push(Promise.resolve().then(() => 'done'))
			} else {
				imageProms.push(createSvgPngPreview(rel))
			}
		})

	return imageProms
}

/**
 * Create SVG preview image
 * @param {ISlideRelMedia} rel - slide rel
 * @return {Promise} promise
 */
async function createSvgPngPreview (rel: ISlideRelMedia): Promise<string> {
	return await new Promise((resolve, reject) => {
		// A: Create
		const image = new Image()

		// B: Set onload event
		image.onload = () => {
			// First: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (image.width + image.height === 0) {
				image.onerror('h/w=0')
			}
			let canvas: HTMLCanvasElement = document.createElement('CANVAS') as HTMLCanvasElement
			const ctx = canvas.getContext('2d')
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
			reject(new Error(`ERROR! Unable to load image (image.onerror): ${rel.path}`))
		}

		// C: Load image
		image.src = typeof rel.data === 'string' ? rel.data : IMG_BROKEN
	})
}

/**
 * FIXME: TODO: currently unused
 * TODO: Should return a Promise
 */
function getSizeFromImage (inImgUrl: string): { width: number, height: number } {
	const sizeOf = typeof require !== 'undefined' ? require('sizeof') : null // NodeJS

	if (sizeOf) {
		try {
			const dimensions = sizeOf(inImgUrl)
			return { width: dimensions.width, height: dimensions.height }
		} catch (ex) {
			console.error('ERROR: sizeOf: Unable to load image: ' + inImgUrl)
			return { width: 0, height: 0 }
		}
	} else if (Image && typeof Image === 'function') {
		// A: Create
		const image = new Image()

		// B: Set onload event
		image.onload = () => {
			// FIRST: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (image.width + image.height === 0) {
				return { width: 0, height: 0 }
			}
			const obj = { width: image.width, height: image.height }
			return obj
		}
		image.onerror = () => {
			console.error(`ERROR: image.onload: Unable to load image: ${inImgUrl}`)
		}

		// C: Load image
		image.src = inImgUrl
	}
}
