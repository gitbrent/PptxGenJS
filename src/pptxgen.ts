/*\
|*|  :: pptxgen.js ::
|*|
|*|  JavaScript framework that creates PowerPoint (pptx) presentations
|*|  https://github.com/gitbrent/PptxGenJS
|*|
|*|  This framework is released under the MIT Public License (MIT)
|*|
|*|  PptxGenJS (C) 2015-2019 Brent Ely -- https://github.com/gitbrent
|*|
|*|  Some code derived from the OfficeGen project:
|*|  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
|*|
|*|  Permission is hereby granted, free of charge, to any person obtaining a copy
|*|  of this software and associated documentation files (the "Software"), to deal
|*|  in the Software without restriction, including without limitation the rights
|*|  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
|*|  copies of the Software, and to permit persons to whom the Software is
|*|  furnished to do so, subject to the following conditions:
|*|
|*|  The above copyright notice and this permission notice shall be included in all
|*|  copies or substantial portions of the Software.
|*|
|*|  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
|*|  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
|*|  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
|*|  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
|*|  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
|*|  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
|*|  SOFTWARE.
\*/

/*
	PPTX Units are "DXA" (except for font sizing)
	....: There are 1440 DXA per inch. 1 inch is 72 points. 1 DXA is 1/20th's of a point (20 DXA is 1 point).
	....: There is also something called EMU's (914400 EMUs is 1 inch, 12700 EMUs is 1pt).
	SEE: https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
	|
	OBJECT LAYOUTS: 16x9 (10" x 5.625"), 16x10 (10" x 6.25"), 4x3 (10" x 7.5"), Wide (13.33" x 7.5") and Custom (any size)
	|
	REFS:
	* "Structure of a PresentationML document (Open XML SDK)"
	* @see: https://msdn.microsoft.com/en-us/library/office/gg278335.aspx
	* TableStyleId enumeration
	* @see: https://msdn.microsoft.com/en-us/library/office/hh273476(v=office.14).aspx
*/

import { CHART_TYPES, DEF_PRES_LAYOUT, DEF_SLIDE_MARGIN_IN, IMG_BROKEN, JSZIP_OUTPUT_TYPE, SCHEME_COLOR_NAMES, SLIDE_OBJECT_TYPES, DEF_PRES_LAYOUT_NAME } from './enums'
import { ISlide, ILayout, ISlideLayout, ISlideRelMedia, ISlideMasterDef, ISlideRel, ISlideNumber, ITableToSlidesOpts } from './interfaces'
import Slide from './pptxgen-slide'
import * as JSZip from 'jszip'
import * as genCharts from './gen-charts'
import * as genObj from './gen-objects'
import * as genXml from './gen-xml'
import { gObjPptxShapes } from './lib-shapes'

export default class PptxGenJS {
	// Property getters/setters

	/**
	 * Presentation Layout: 'screen4x3', 'screen16x9', 'widescreen', etc.
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 */
	private _layout: string
	public set layout(value: string) {
		let newLayout: ILayout = this.LAYOUTS[value]

		if (newLayout) {
			this._layout = value
			this._presLayout = newLayout
		} else {
			throw 'UNKNOWN-LAYOUT'
		}
	}
	public get layout(): string {
		return this._layout
	}

	private _version: string = '3.0.0-beta1'
	public get version(): string {
		return this._version
	}

	private _author: string
	public set author(value: string) {
		this._author = value
	}
	public get author(): string {
		return this._author
	}

	private _company: string
	public set company(value: string) {
		this._company = value
	}
	public get company(): string {
		return this._company
	}

	/**
	 * Sets the Presentation's Revision
	 * PowerPoint requires `revision` be a number only (without "." or ",") (otherwise, PPT will throw errors upon opening Presentation!)
	 */
	private _revision: string
	public set revision(value: string) {
		this._revision = value
	}
	public get revision(): string {
		return this._revision
	}

	private _subject: string
	public set subject(value: string) {
		this._subject = value
	}
	public get subject(): string {
		return this._subject
	}

	private _title: string
	public set title(value: string) {
		this._title = value
	}
	public get title(): string {
		return this._title
	}

	/**
	 * Whether Right-to-Left (RTL) mode is enabled
	 */
	private _rtlMode: boolean
	public set rtlMode(value: boolean) {
		this._rtlMode = value
	}
	public get rtlMode(): boolean {
		return this._rtlMode
	}

	/**
	 * `isBrowser` Presentation Option:
	 * Target: Angular/React/Webpack, etc. This setting affects how files are saved: using `fs` for Node.js or browser libs
	 */
	private _isBrowser: boolean
	public set isBrowser(value: boolean) {
		this._isBrowser = value
	}
	public get isBrowser(): boolean {
		return this._isBrowser
	}

	// TODO: should these be `this.var` inside constructor?
	private fileName: string
	private fileExtn: string
	/** master slide layout object */
	private masterSlide: ISlide

	/** this Presentation's Slide objects */
	private slides: ISlide[]

	/** slide layout definition objects, used for generating slide layout files */
	private slideLayouts: ISlideLayout[]
	private saveCallback: Function
	private NODEJS: boolean = false
	private LAYOUTS: object

	// Global props
	private _charts = CHART_TYPES
	public get charts(): typeof CHART_TYPES {
		return this._charts
	}
	private _colors = SCHEME_COLOR_NAMES
	public get colors(): typeof SCHEME_COLOR_NAMES {
		return this._colors
	}
	private _shapes = gObjPptxShapes
	public get shapes(): typeof gObjPptxShapes {
		return this._shapes
	}

	private _presLayout: ILayout
	presLayout = (): ILayout => {
		return this._presLayout
	}

	//private _imageCounter: number // TODO: This is a dummy value - `gen-xml` has real one: find a better solution, stop using counter

	private fs: any
	private https: any
	private sizeOf: any

	constructor() {
		// Determine Environment
		if (typeof module !== 'undefined' && module.exports && typeof require === 'function' && typeof window === 'undefined') {
			try {
				require.resolve('fs')
				this.NODEJS = true
			} catch (ex) {
				this.NODEJS = false
			}
		}

		// Set available layouts
		this.LAYOUTS = {
			LAYOUT_4x3: { name: 'screen4x3', width: 9144000, height: 6858000 } as ILayout,
			LAYOUT_16x9: { name: 'screen16x9', width: 9144000, height: 5143500 } as ILayout,
			LAYOUT_16x10: { name: 'screen16x10', width: 9144000, height: 5715000 } as ILayout,
			LAYOUT_WIDE: { name: 'custom', width: 12192000, height: 6858000 } as ILayout,
			LAYOUT_USER: { name: 'custom', width: 12192000, height: 6858000 } as ILayout,
		}

		// Core
		this._author = 'PptxGenJS'
		this._company = 'PptxGenJS'
		this._revision = '1' // whole number
		this._subject = 'PptxGenJS Presentation'
		this._title = 'PptxGenJS Presentation'
		// PptxGenJS props
		this._presLayout = {
			name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
			width: this.LAYOUTS[DEF_PRES_LAYOUT].width,
			height: this.LAYOUTS[DEF_PRES_LAYOUT].height,
		}
		this._rtlMode = false
		this._isBrowser = false
		this.fileName = 'Presentation'
		this.fileExtn = '.pptx'
		//this.saveCallback = null // FIXME: deprecated: moving to Promise
		//
		this.slideLayouts = [
			{
				presLayout: this._presLayout,
				name: DEF_PRES_LAYOUT_NAME,
				width: this.LAYOUTS['LAYOUT_16x9'].width,
				height: this.LAYOUTS['LAYOUT_16x9'].height,
				slide: null,
				data: [],
				rels: [],
				relsChart: [],
				relsMedia: [],
				margin: DEF_SLIDE_MARGIN_IN,
				slideNumberObj: null,
			},
		]
		this.slides = []

		this.masterSlide = {
			addChart: null,
			addImage: null,
			addMedia: null,
			addNotes: null,
			addShape: null,
			addTable: null,
			addText: null,
			//
			presLayout: this._presLayout,
			name: null,
			number: null,
			data: [],
			rels: [],
			relsChart: [],
			relsMedia: [],
			slideLayout: null,
			slideNumberObj: null,
		}
	}

	/**
	 * Create and export the .pptx file
	 * @param {JSZIP_OUTPUT_TYPE} `outputType` JSZip output type (ArrayBuffer, Blob, etc.)
	 */
	doExportPresentation = (outputType?: JSZIP_OUTPUT_TYPE) => {
		var arrChartPromises: Array<Promise<any>> = []
		var intSlideNum: number = 0

		// STEP 1: Create new JSZip file
		let zip: JSZip = new JSZip()

		// STEP 2: Add all required folders and files
		zip.folder('_rels')
		zip.folder('docProps')
		zip.folder('ppt').folder('_rels')
		zip.folder('ppt/charts').folder('_rels')
		zip.folder('ppt/embeddings')
		zip.folder('ppt/media')
		zip.folder('ppt/slideLayouts').folder('_rels')
		zip.folder('ppt/slideMasters').folder('_rels')
		zip.folder('ppt/slides').folder('_rels')
		zip.folder('ppt/theme')
		zip.folder('ppt/notesMasters').folder('_rels')
		zip.folder('ppt/notesSlides').folder('_rels')
		zip.file('[Content_Types].xml', genXml.makeXmlContTypes(this.slides, this.slideLayouts, this.masterSlide))
		zip.file('_rels/.rels', genXml.makeXmlRootRels())
		zip.file('docProps/app.xml', genXml.makeXmlApp(this.slides, this.company))
		zip.file('docProps/core.xml', genXml.makeXmlCore(this.title, this.subject, this.author, this.revision))
		zip.file('ppt/_rels/presentation.xml.rels', genXml.makeXmlPresentationRels(this.slides))
		zip.file('ppt/theme/theme1.xml', genXml.makeXmlTheme())
		zip.file('ppt/presentation.xml', genXml.makeXmlPresentation(this.slides, this._presLayout))
		zip.file('ppt/presProps.xml', genXml.makeXmlPresProps())
		zip.file('ppt/tableStyles.xml', genXml.makeXmlTableStyles())
		zip.file('ppt/viewProps.xml', genXml.makeXmlViewProps())

		// STEP 3: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
		for (var idx = 1; idx <= this.slideLayouts.length; idx++) {
			zip.file('ppt/slideLayouts/slideLayout' + idx + '.xml', genXml.makeXmlLayout(this.slideLayouts[idx - 1]))
			zip.file('ppt/slideLayouts/_rels/slideLayout' + idx + '.xml.rels', genXml.makeXmlSlideLayoutRel(idx, this.slideLayouts))
		}
		for (var idx = 0; idx < this.slides.length; idx++) {
			intSlideNum++
			zip.file('ppt/slides/slide' + intSlideNum + '.xml', genXml.makeXmlSlide(this.slides[idx]))
			zip.file('ppt/slides/_rels/slide' + intSlideNum + '.xml.rels', genXml.makeXmlSlideRel(this.slides, this.slideLayouts, intSlideNum))

			// Here we will create all slide notes related items. Notes of empty strings
			// are created for slides which do not have notes specified, to keep track of _rels.
			zip.file('ppt/notesSlides/notesSlide' + intSlideNum + '.xml', genXml.makeXmlNotesSlide(this.slides[idx]))
			zip.file('ppt/notesSlides/_rels/notesSlide' + intSlideNum + '.xml.rels', genXml.makeXmlNotesSlideRel(intSlideNum))
		}
		zip.file('ppt/slideMasters/slideMaster1.xml', genXml.makeXmlMaster(this.masterSlide, this.slideLayouts))
		zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', genXml.makeXmlMasterRel(this.masterSlide, this.slideLayouts))
		zip.file('ppt/notesMasters/notesMaster1.xml', genXml.makeXmlNotesMaster())
		zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', genXml.makeXmlNotesMasterRel())

		// STEP 4: Create all Rels (images, media, chart data)
		this.slideLayouts.forEach(layout => {
			this.createChartMediaRels(layout, zip, arrChartPromises)
		})
		this.slides.forEach(slide => {
			this.createChartMediaRels(slide, zip, arrChartPromises)
		})
		this.createChartMediaRels(this.masterSlide, zip, arrChartPromises)

		// STEP 5: Wait for Promises (if any) then generate the PPTX file
		Promise.all(arrChartPromises)
			.then(() => {
				var strExportName = this.fileName.toLowerCase().indexOf('.ppt') > -1 ? this.fileName : this.fileName + this.fileExtn
				if (outputType) {
					zip.generateAsync({ type: outputType }).then(() => this.saveCallback)
				} else if (this.NODEJS && !this.isBrowser) {
					if (this.saveCallback) {
						if (strExportName.indexOf('http') == 0) {
							zip.generateAsync({ type: 'nodebuffer' }).then(content => {
								this.saveCallback(content)
							})
						} else {
							zip.generateAsync({ type: 'nodebuffer' }).then(content => {
								this.fs.writeFile(strExportName, content, () => {
									this.saveCallback(strExportName)
								})
							})
						}
					} else {
						// Starting in late 2017 (Node ~8.9.1), `fs` requires a callback so use a dummy func
						zip.generateAsync({ type: 'nodebuffer' }).then(content => {
							this.fs.writeFile(strExportName, content, () => {})
						})
					}
				} else {
					zip.generateAsync({ type: 'blob' }).then(content => {
						this.writeFileToBrowser(strExportName, content)
					})
				}
			})
			.catch(strErr => {
				console.error(strErr)
			})
	}

	writeFileToBrowser = (strExportName: string, content: Blob) => {
		// STEP 1: Create element
		let a = document.createElement('a')
		a.setAttribute('style', 'display:none;')
		document.body.appendChild(a)

		// STEP 2: Download file to browser
		// DESIGN: Use `createObjectURL()` (or MS-specific func for IE11) to D/L files in client browsers (FYI: synchronously executed)
		if (window.navigator.msSaveOrOpenBlob) {
			// REF: https://docs.microsoft.com/en-us/microsoft-edge/dev-guide/html5/file-api/blob
			let blobObject = new Blob([content])
			jQuery(a).click(() => {
				window.navigator.msSaveOrOpenBlob(blobObject, strExportName)
			})
			a.click()

			// Clean-up
			document.body.removeChild(a)

			// LAST: Callback (if any)
			if (this.saveCallback) this.saveCallback(strExportName)
		} else if (window.URL.createObjectURL) {
			var blob = new Blob([content], { type: 'octet/stream' })
			var url = window.URL.createObjectURL(blob)
			a.href = url
			a.download = strExportName
			a.click()

			// Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
			setTimeout(() => {
				window.URL.revokeObjectURL(url)
				document.body.removeChild(a)
			}, 100)

			// LAST: Callback (if any)
			if (this.saveCallback) this.saveCallback(strExportName)
		}

		// STEP 3: Clear callback func post-save
		this.saveCallback = null
	}

	createChartMediaRels = (slide: ISlide | ISlideLayout, zip: JSZip, chartPromises: Array<Promise<any>>) => {
		slide.relsChart.forEach(rel => chartPromises.push(genCharts.createExcelWorksheet(rel, zip)))
		slide.relsMedia.forEach(rel => {
			if (rel.type != 'online' && rel.type != 'hyperlink') {
				// A: Loop vars
				var data: string = rel.data as string

				// B: Users will undoubtedly pass various string formats, so correct prefixes as needed
				if (data.indexOf(',') == -1 && data.indexOf(';') == -1) data = 'image/png;base64,' + data
				else if (data.indexOf(',') == -1) data = 'image/png;base64,' + data
				else if (data.indexOf(';') == -1) data = 'image/png;' + data

				// C: Add media
				zip.file(rel.Target.replace('..', 'ppt'), data.split(',').pop(), { base64: true })
			}
		})
	}

	addPlaceholdersToSlides = (slide: ISlide) => {
		// Add all placeholders on this Slide that dont already exist
		;(slide.slideLayout.data || []).forEach(slideLayoutObj => {
			if (slideLayoutObj.type === SLIDE_OBJECT_TYPES.placeholder) {
				// A: Search for this placeholder on Slide before we add
				// NOTE: Check to ensure a placeholder does not already exist on the Slide
				// They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
				if (
					slide.data.filter(slideObj => {
						return slideObj.options && slideObj.options.placeholder == slideLayoutObj.options.placeholder
					}).length == 0
				) {
					genObj.addTextDefinition('', { placeholder: slideLayoutObj.options.placeholder }, slide, false)
				}
			}
		})
	}

	// IMAGE METHODS:

	getSizeFromImage = (inImgUrl: string) => {
		if (this.NODEJS) {
			try {
				var dimensions = this.sizeOf(inImgUrl)
				return { width: dimensions.width, height: dimensions.height }
			} catch (ex) {
				console.error('ERROR: Unable to read image: ' + inImgUrl)
				return { width: 0, height: 0 }
			}
		}

		// A: Create
		var image = new Image()

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
				console.error('[Error] Unable to load image: ' + inImgUrl)
			} catch (ex) {}
		}

		// C: Load image
		image.src = inImgUrl
	}

	/* Encode Image/Audio/Video into base64 */
	encodeSlideMediaRels = (layout, arrRelsDone) => {
		let intRels = 0

		layout.relsMedia.forEach((rel: ISlideRelMedia) => {
			// Read and Encode each media lacking `data` into base64 (for use in export)
			if (rel.type != 'online' && !rel.data && arrRelsDone.indexOf(rel.path) == -1) {
				// Node local-file encoding is syncronous, so we can load all images here, then call export with a callback (if any)
				if (this.NODEJS && rel.path.indexOf('http') != 0) {
					try {
						var bitmap = this.fs.readFileSync(rel.path)
						rel.data = Buffer.from(bitmap).toString('base64')
					} catch (ex) {
						console.error('ERROR....: Unable to read media: "' + rel.path + '"')
						console.error('DETAILS..: ' + ex)
						rel.data = IMG_BROKEN
					}
				} else if (this.NODEJS && rel.path.indexOf('http') == 0) {
					intRels++
					this.convertRemoteMediaToDataURL(rel)
					arrRelsDone.push(rel.path)
				} else {
					intRels++
					this.convertImgToDataURL(rel)
					arrRelsDone.push(rel.path)
				}
			} else if (
				rel.isSvgPng &&
				rel.data &&
				rel.data
					.toString()
					.toLowerCase()
					.indexOf('image/svg') > -1
			) {
				// The SVG base64 must be converted to PNG SVG before export
				intRels++
				this.callbackImgToDataURLDone(rel.data, rel)
				arrRelsDone.push(rel.path)
			}
		})

		return intRels
	}

	/* `FileReader()` + `readAsDataURL` = Ablity to read any file into base64! */
	convertImgToDataURL = (slideRel: ISlideRelMedia) => {
		var xhr = new XMLHttpRequest()
		xhr.onload = () => {
			var reader = new FileReader()
			reader.onloadend = () => {
				this.callbackImgToDataURLDone(reader.result, slideRel)
			}
			reader.readAsDataURL(xhr.response)
		}
		xhr.onerror = ex => {
			// TODO: xhr.error/catch whatever! then return
			console.error('Unable to load image: "' + slideRel.path)
			console.error(ex || '')
			// Return a predefined "Broken image" graphic so the user will see something on the slide
			this.callbackImgToDataURLDone(IMG_BROKEN, slideRel)
		}
		xhr.open('GET', slideRel.path)
		xhr.responseType = 'blob'
		xhr.send()
	}

	/**
	 * Node equivalent of `convertImgToDataURL()`: Use https to fetch, then use Buffer to encode to base64
	 * @param {ISlideRelMedia} `slideRel` - slide rel
	 */
	convertRemoteMediaToDataURL = (slideRel: ISlideRelMedia) => {
		this.https.get(slideRel.path, res => {
			var rawData = ''
			res.setEncoding('binary') // IMPORTANT: Only binary encoding works
			res.on('data', chunk => {
				rawData += chunk
			})
			res.on('end', () => {
				var data = Buffer.from(rawData, 'binary').toString('base64')
				this.callbackImgToDataURLDone(data, slideRel)
			})
			res.on('error', e => {
				// TODO-3: make this method return Promise?
				///reject(e);
			})
		})
	}

	/**
	 * (Browser Only): Convert SVG-base64 data to PNG-base64
	 * @param {ISlideRelMedia} `slideRel` - slide rel
	 */
	convertSvgToPngViaCanvas = (slideRel: ISlideRelMedia) => {
		// A: Create
		let image = new Image()

		// B: Set onload event
		image.onload = () => {
			// First: Check for any errors: This is the best method (try/catch wont work, etc.)
			if (image.width + image.height == 0) {
				image.onerror('h/w=0')
				return
			}
			var canvas: HTMLCanvasElement = document.createElement('CANVAS') as HTMLCanvasElement
			var ctx = canvas.getContext('2d')
			canvas.width = image.width
			canvas.height = image.height
			ctx.drawImage(image, 0, 0)
			// Users running on local machine will get the following error:
			// "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
			// when the canvas.toDataURL call executes below.
			try {
				this.callbackImgToDataURLDone(canvas.toDataURL(slideRel.type), slideRel)
			} catch (ex) {
				image.onerror(ex)
				return
			}
			canvas = null
		}
		image.onerror = ex => {
			console.error(ex || '')
			// Return a predefined "Broken image" graphic so the user will see something on the slide
			this.callbackImgToDataURLDone(IMG_BROKEN, slideRel)
		}

		// C: Load image
		image.src = slideRel.data as string // use pre-encoded SVG base64 data
	}

	// FIXME: 20190715: if all were doing here is mapping `rel.base64` to `rel.data`, surely there's a better/clearer way...
	callbackImgToDataURLDone = (base64Data: string | ArrayBuffer, slideRel: ISlideRelMedia) => {
		// SVG images were retrieved via `convertImgToDataURL()`, but have to be encoded to PNG now
		if (slideRel.isSvgPng && typeof base64Data === 'string' && base64Data.indexOf('image/svg') > -1) {
			// Pass the SVG XML as base64 for conversion to PNG

			slideRel.data = base64Data
			if (this.NODEJS) throw 'SVG is not supported in Node (more info: https://github.com/gitbrent/PptxGenJS/issues/401)'
			else this.convertSvgToPngViaCanvas(slideRel)
			return
		}

		let intEmpty = 0
		let funcCallback = (rel: ISlideRel | ISlideRelMedia) => {
			if (rel.path == slideRel.path) rel.data = base64Data
			if (!rel.data) intEmpty++
		}

		// STEP 1: Set data for this rel, count outstanding
		this.slides.forEach(slide => {
			slide.rels.forEach(funcCallback)
			slide.relsMedia.forEach(funcCallback)
		})
		this.slideLayouts.forEach(layout => {
			layout.rels.forEach(funcCallback)
			layout.relsMedia.forEach(funcCallback)
		})
		this.masterSlide.rels.forEach(funcCallback)
		this.masterSlide.relsMedia.forEach(funcCallback)

		// STEP 2: Continue export process if all rels have base64 `data` now
		if (intEmpty == 0) this.doExportPresentation()
	}

	/**
	 * Enables the Slide class to set PptxGen master/layout slidenumbers
	 */
	setSlideNumber = (slideNumberObj: ISlideNumber) => {
		// 1: Add slideNumber to slideMaster1.xml
		this.masterSlide.slideNumberObj = slideNumberObj

		// 2: Add slideNumber to DEF_PRES_LAYOUT_NAME layout
		this.slideLayouts.filter(layout => {
			return layout.name == DEF_PRES_LAYOUT_NAME
		})[0].slideNumberObj = slideNumberObj
	}

	// PUBLIC API

	/**
	 * Save (export) the Presentation .pptx file
	 * @param {string} `exportName` - Filename to use for the export
	 * @param {Function} `callbackFunc` - Callback function to be called when export is complete
	 * @param {JSZIP_OUTPUT_TYPE} `outputType` - JSZip output type
	 */
	save(exportName: string, callbackFunc?: Function, outputType?: JSZIP_OUTPUT_TYPE) {
		let intRels = 0,
			arrRelsDone = []

		// STEP 1: Add empty placeholder objects to slides that don't already have them
		this.slides.forEach(slide => {
			if (slide.slideLayout) this.addPlaceholdersToSlides(slide)
		})

		// STEP 2: Set export properties
		if (callbackFunc) this.saveCallback = callbackFunc
		if (exportName) this.fileName = exportName

		// STEP 3: Read/Encode Images
		// PERF: Only send unique paths for encoding (encoding func will find and fill *ALL* matching paths across the Presentation)

		// A: Slide rels
		this.slides.forEach(slide => {
			intRels += this.encodeSlideMediaRels(slide, arrRelsDone)
		})

		// B: Layout rels
		this.slideLayouts.forEach(layout => {
			intRels += this.encodeSlideMediaRels(layout, arrRelsDone)
		})

		// C: Master Slide rels
		intRels += this.encodeSlideMediaRels(this.masterSlide, arrRelsDone)

		// STEP 4: Export now if there's no images to encode (otherwise, last async imgConvert call above will call exportFile)
		if (intRels == 0) this.doExportPresentation(outputType)
	}

	/**
	 * Add a Slide to Presenation
	 * @param {string} `masterSlideName` Master Slide name
	 * @returns {ISlide} the new Slide
	 */
	addSlide(masterSlideName?: string): ISlide {
		let newSlide = new Slide({
			presLayout: this._presLayout,
			setSlideNum: this.setSlideNumber,
			slideNumber: this.slides.length + 1,
			slideLayout: masterSlideName
				? this.slideLayouts.filter(layout => {
						return layout.name == masterSlideName
				  })[0] || this.LAYOUTS[DEF_PRES_LAYOUT]
				: this.LAYOUTS[DEF_PRES_LAYOUT],
		})

		this.slides.push(newSlide)

		return newSlide
	}

	/**
	 * Adds a new slide master [layout] to the presentation.
	 * @param {ISlideMasterDef} inObjMasterDef - layout definition
	 */
	defineSlideMaster(inObjMasterDef: ISlideMasterDef) {
		if (!inObjMasterDef.title) {
			throw Error('defineSlideMaster() object argument requires a `title` value.')
		}

		let objLayout: ISlideLayout = {
			presLayout: this._presLayout,
			name: inObjMasterDef.title,
			width: inObjMasterDef.width || this._presLayout.width,
			height: inObjMasterDef.height || this._presLayout.height,
			slide: null,
			data: [],
			rels: [],
			relsChart: [],
			relsMedia: [],
			margin: inObjMasterDef.margin || DEF_SLIDE_MARGIN_IN,
			slideNumberObj: inObjMasterDef.slideNumber || null,
		}

		// STEP 1: Create the Slide Master/Layout
		genObj.createSlideObject(inObjMasterDef, objLayout)

		// STEP 2: Add it to layout defs
		this.slideLayouts.push(objLayout)

		// STEP 3: Add slideNumber to master slide (if any)
		if (objLayout.slideNumberObj && !this.masterSlide.slideNumberObj) this.masterSlide.slideNumberObj = objLayout.slideNumberObj
	}

	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} `tabEleId` - HTMLElementID of the table
	 * @param {ITableToSlidesOpts} `inOpts` - array of options (e.g.: tabsize)
	 */
	tableToSlides(tableElementId: string, opts: ITableToSlidesOpts) {
		genXml.genTableToSlides(
			this,
			tableElementId,
			opts,
			opts.masterSlideName
				? this.slideLayouts.filter(layout => {
						return layout.name == opts.masterSlideName.toString()
				  })[0]
				: null
		)
	}
}

/*
// NodeJS support
if (this.NODEJS) {
	jQuery = null
	fs = null
	https = null
	JSZip = null
	sizeOf = null

	// A: jQuery dependency
	try {
		var jsdom = require('jsdom')
		var dom = new jsdom.JSDOM('<!DOCTYPE html>')
		jQuery = require('jquery')(dom.window)
	} catch (ex) {
		console.error('Unable to load `jquery`!\n' + ex)
		throw 'LIB-MISSING-JQUERY'
	}

	// B: Other dependencies
	try {
		fs = require('fs')
	} catch (ex) {
		console.error('Unable to load `fs`')
		throw 'LIB-MISSING-FS'
	}
	try {
		https = require('https')
	} catch (ex) {
		console.error('Unable to load `https`')
		throw 'LIB-MISSING-HTTPS'
	}
	try {
		JSZip = require('jszip')
	} catch (ex) {
		console.error('Unable to load `jszip`')
		throw 'LIB-MISSING-JSZIP'
	}
	try {
		sizeOf = require('image-size')
	} catch (ex) {
		console.error('Unable to load `image-size`')
		throw 'LIB-MISSING-IMGSIZE'
	}

	// LAST: Export module
	module.exports = PptxGenJS
}
// Angular/React/etc support
else if (APPJS) {
	// A: jQuery dependency
	try {
		jQuery = require('jquery')
	} catch (ex) {
		console.error('Unable to load `jquery`!\n' + ex)
		throw 'LIB-MISSING-JQUERY'
	}

	// B: Other dependencies
	try {
		JSZip = require('jszip')
	} catch (ex) {
		console.error('Unable to load `jszip`')
		throw 'LIB-MISSING-JSZIP'
	}

	// LAST: Export module
	module.exports = PptxGenJS
}
*/
