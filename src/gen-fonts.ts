import { PresFont } from './core-interfaces'
import { IMG_BROKEN } from './core-enums'

export function encodePresFontRels(fonts: PresFont[]): Promise<string>[] {
	let promises: Promise<string>[] = []

	// A: Capture all audio/image/video candidates for encoding (filtering online/pre-encoded)
	let candidateRels = fonts
		.map(({ variants }) => variants)
		.flat()
		.filter(variant => !variant.data && (!variant.path || (variant.path && variant.path.indexOf('preencoded') === -1)))

	// B: PERF: Mark dupes (same `path`) so that we dont load same media over-and-over
	let unqPaths: string[] = []
	candidateRels.forEach(rel => {
		if (unqPaths.indexOf(rel.path) === -1) {
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
			promises.push(
				new Promise((resolve, reject) => {
					// A: Declare XHR and onload/onerror handlers
					// DESIGN: `XMLHttpRequest()` plus `FileReader()` = Ablity to read any file into base64!
					let xhr = new XMLHttpRequest()
					xhr.onload = () => {
						let reader = new FileReader()
						reader.onloadend = () => {
							console.log('font', rel.type, reader.result)
							rel.data = reader.result
							candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
							resolve('done')
						}
						reader.readAsDataURL(xhr.response)
					}

					xhr.onerror = ex => {
						rel.data = IMG_BROKEN
						candidateRels.filter(dupe => dupe.isDuplicate && dupe.path === rel.path).forEach(dupe => (dupe.data = rel.data))
						reject(`ERROR! Unable to load image (xhr.onerror): ${rel.path}`)
					}

					// B: Execute request
					xhr.open('GET', rel.path)
					xhr.responseType = 'blob'
					xhr.send()
				})
			)
		})

	return promises
}
