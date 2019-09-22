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

/**
 * PPTX Units are "DXA" (except for font sizing)
 * ....: There are 1440 DXA per inch. 1 inch is 72 points. 1 DXA is 1/20th's of a point (20 DXA is 1 point).
 * ....: There is also something called EMU's (914400 EMUs is 1 inch, 12700 EMUs is 1pt).
 * SEE: https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
 *
 * OBJECT LAYOUTS: 16x9 (10" x 5.625"), 16x10 (10" x 6.25"), 4x3 (10" x 7.5"), Wide (13.33" x 7.5") and Custom (any size)
 *
 * REFERENCES:
 * @see [Structure of a PresentationML document (Open XML SDK)](https://msdn.microsoft.com/en-us/library/office/gg278335.aspx)
 * @see [TableStyleId enumeration](https://msdn.microsoft.com/en-us/library/office/hh273476(v=office.14).aspx)
 */

import { CHART_TYPES, DEF_PRES_LAYOUT_NAME, DEF_PRES_LAYOUT, DEF_SLIDE_MARGIN_IN, JSZIP_OUTPUT_TYPE, SCHEME_COLOR_NAMES, WRITE_OUTPUT_TYPE } from './core-enums'
import { ILayout, ISlide, ISlideLayout, ISlideMasterOptions, ISlideNumber, ITableToSlidesOpts } from './core-interfaces'
import { PowerPointShapes } from './core-shapes'
import Slide from './slide'
import * as genCharts from './gen-charts'
import * as genObj from './gen-objects'
import * as genMedia from './gen-media'
import * as genTable from './gen-tables'
import * as genXml from './gen-xml'
import * as JSZip from 'jszip'

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

	private _version: string = '3.0.0-beta2'
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

	/** master slide layout object */
	private masterSlide: ISlide

	/** this Presentation's Slide objects */
	private slides: ISlide[]

	/** slide layout definition objects, used for generating slide layout files */
	private slideLayouts: ISlideLayout[]
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
	private _shapes = PowerPointShapes
	public get shapes(): typeof PowerPointShapes {
		return this._shapes
	}
	private _presLayout: ILayout
	public get presLayout(): ILayout {
		return this._presLayout
	}

	constructor() {
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
		this._revision = '1' // Note: Must be a whole number
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
		//
		this.slideLayouts = [
			{
				presLayout: this._presLayout,
				name: DEF_PRES_LAYOUT_NAME,
				number: 1000,
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
	 * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
	 * @param {string} masterName - slide master name
	 * @return {ISlide} new Slide
	 */
	addNewSlide = (masterName: string): ISlide => {
		return this.addSlide(masterName)
	}

	/**
	 * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
	 * @param {number} slideNum - slide number
	 * @return {ISlide} Slide
	 * @since 3.0.0
	 */
	getSlide = (slideNum: number): ISlide => {
		return this.slides.filter(slide => {
			return slide.number == slideNum
		})[0]
	}

	/**
	 * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
	 * @param {ISlideNumber} slideNum - slide number config
	 */
	setSlideNumber = (slideNum: ISlideNumber) => {
		// 1: Add slideNumber to slideMaster1.xml
		this.masterSlide.slideNumberObj = slideNum

		// 2: Add slideNumber to DEF_PRES_LAYOUT_NAME layout
		this.slideLayouts.filter(layout => {
			return layout.name == DEF_PRES_LAYOUT_NAME
		})[0].slideNumberObj = slideNum
	}

	/**
	 * Create all chart and media rels for this Presenation
	 * @param {ISlide | ISlideLayout} slide - slide with rels
	 * @param {JSZIP} zip - JSZip instance
	 * @param {Promise<any>[]} chartPromises - promise array
	 */
	createChartMediaRels = (slide: ISlide | ISlideLayout, zip: JSZip, chartPromises: Promise<any>[]) => {
		slide.relsChart.forEach(rel => chartPromises.push(genCharts.createExcelWorksheet(rel, zip)))
		slide.relsMedia.forEach(rel => {
			if (rel.type != 'online' && rel.type != 'hyperlink') {
				// A: Loop vars
				let data: string = rel.data && typeof rel.data === 'string' ? rel.data : ''

				// B: Users will undoubtedly pass various string formats, so correct prefixes as needed
				if (data.indexOf(',') == -1 && data.indexOf(';') == -1) data = 'image/png;base64,' + data
				else if (data.indexOf(',') == -1) data = 'image/png;base64,' + data
				else if (data.indexOf(';') == -1) data = 'image/png;' + data

				// C: Add media
				zip.file(rel.Target.replace('..', 'ppt'), data.split(',').pop(), { base64: true })
			}
		})
	}

	/**
	 * Create and export the .pptx file
	 * @param {WRITE_OUTPUT_TYPE} outputType - output file type
	 * @return {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} Promise with data or stream (node) or filename (browser)
	 */
	exportPresentation = (outputType?: WRITE_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> => {
		return new Promise((resolve, reject) => {
			let arrChartPromises: Promise<string>[] = []
			let arrMediaPromises: Promise<string>[] = []
			let zip: JSZip = new JSZip()

			// STEP 1: Read/Encode all Media before zip as base64 content, etc. is required
			this.slides.forEach(slide => {
				arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(slide))
			})
			this.slideLayouts.forEach(layout => {
				arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(layout))
			})
			arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(this.masterSlide))

			// STEP 6: Wait for Promises (if any) then generate the PPTX file
			Promise.all(arrMediaPromises).then(() => {
				// STEP 1: Add empty placeholder objects to slides that don't already have them
				this.slides.forEach(slide => {
					if (slide.slideLayout) genObj.addPlaceholdersToSlideLayouts(slide)
				})

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
				zip.file('ppt/presentation.xml', genXml.makeXmlPresentation(this.slides, this.presLayout, this.rtlMode))
				zip.file('ppt/presProps.xml', genXml.makeXmlPresProps())
				zip.file('ppt/tableStyles.xml', genXml.makeXmlTableStyles())
				zip.file('ppt/viewProps.xml', genXml.makeXmlViewProps())

				// STEP 3: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
				this.slideLayouts.forEach((layout, idx) => {
					zip.file('ppt/slideLayouts/slideLayout' + (idx + 1) + '.xml', genXml.makeXmlLayout(layout))
					zip.file('ppt/slideLayouts/_rels/slideLayout' + (idx + 1) + '.xml.rels', genXml.makeXmlSlideLayoutRel(idx + 1, this.slideLayouts))
				})
				this.slides.forEach((slide, idx) => {
					zip.file('ppt/slides/slide' + (idx + 1) + '.xml', genXml.makeXmlSlide(slide))
					zip.file('ppt/slides/_rels/slide' + (idx + 1) + '.xml.rels', genXml.makeXmlSlideRel(this.slides, this.slideLayouts, idx + 1))
					// Create all slide notes related items. Notes of empty strings are created for slides which do not have notes specified, to keep track of _rels.
					zip.file('ppt/notesSlides/notesSlide' + (idx + 1) + '.xml', genXml.makeXmlNotesSlide(slide))
					zip.file('ppt/notesSlides/_rels/notesSlide' + (idx + 1) + '.xml.rels', genXml.makeXmlNotesSlideRel(idx + 1))
				})
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

				// STEP 6: Wait for Promises (if any) then generate the PPTX file
				Promise.all(arrChartPromises)
					.then(() => {
						if (outputType == 'STREAM') {
							// A: stream file
							zip.generateAsync({ type: 'nodebuffer' }).then(content => { resolve(content) })
						} else if (outputType) {
							// B: Node [fs]: Output type user option or default
							resolve(zip.generateAsync({ type: outputType || 'nodebuffer' }))
						} else {
							// C: Browser: Output blob as app/ms-pptx
							resolve(zip.generateAsync({ type: 'blob' }))
						}
					})
					.catch(err => {
						reject(err)
					})
			})
		})
	}

	/**
	 * Create and export the .pptx file
	 * @param {string} exportName - output file type
	 * @param {Blob} content - output file type
	 * @return {Promise<string>} Promise with file name
	 */
	writeFileToBrowser = (exportName: string, blobContent: Blob): Promise<string> => {
		return new Promise((resolve, _reject) => {
			// STEP 1: Create element
			let eleLink = document.createElement('a')
			eleLink.setAttribute('style', 'display:none;')
			document.body.appendChild(eleLink)

			// STEP 2: Download file to browser
			// DESIGN: Use `createObjectURL()` (or MS-specific func for IE11) to D/L files in client browsers (FYI: synchronously executed)
			if (window.navigator.msSaveOrOpenBlob) {
				// @see https://docs.microsoft.com/en-us/microsoft-edge/dev-guide/html5/file-api/blob
				let blob = new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
				eleLink.onclick = function() {
					window.navigator.msSaveOrOpenBlob(blob, exportName)
				}
				eleLink.click()

				// Clean-up
				document.body.removeChild(eleLink)

				// Done
				resolve(exportName)
			} else if (window.URL.createObjectURL) {
				let url = window.URL.createObjectURL(new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }))
				eleLink.href = url
				eleLink.download = exportName
				eleLink.click()

				// Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
				setTimeout(() => {
					window.URL.revokeObjectURL(url)
					document.body.removeChild(eleLink)
				}, 100)

				// Done
				resolve(exportName)
			}
		})
	}

	// EXPORT METHODS

	/**
	 * Export the current Presenation to stream
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
	 * @since 3.0.0
	 */
	stream(): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> {
		return new Promise((resolve, reject) => {
			this.exportPresentation('STREAM')
				.then(content => {
					resolve(content)
				})
				.catch(ex => {
					reject(ex)
				})
		})
	}

	/**
	 * Export the current Presenation to selected/default type
	 * @param {JSZIP_OUTPUT_TYPE} outputType - Master Slide name
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file
	 * @since 3.0.0
	 */
	write(outputType: JSZIP_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> {
		return new Promise((resolve, reject) => {
			this.exportPresentation(outputType)
				.then(content => {
					resolve(content)
				})
				.catch(ex => {
					reject(ex + '\nDid you mean to use writeFile() instead?')
				})
		})
	}

	/**
	 * Export the current Presenation to local file (initiates download in browsers)
	 * @param {string} exportName - file name
	 * @returns {Promise<string>} file name
	 * @since 3.0.0
	 */
	writeFile(exportName?: string): Promise<string> {
		return new Promise((resolve, reject) => {
			const fs = typeof require !== 'undefined' ? require('fs') : null // NodeJS
			let fileName = exportName
				? exportName
						.toString()
						.toLowerCase()
						.endsWith('.pptx')
					? exportName
					: exportName + '.pptx'
				: 'Presenation.pptx'

			this.exportPresentation()
				.then(content => {
					if (fs) {
						// Node: Output
						fs.writeFile(fileName, content, () => {
							resolve(fileName)
						})
					} else {
						// Browser: Output blob as app/ms-pptx
						resolve(this.writeFileToBrowser(fileName, content as Blob))
					}
				})
				.catch(ex => {
					reject(ex)
				})
		})
	}

	// PRESENTATION METHODS

	/**
	 * Add a Slide to Presenation
	 * @param {string} masterSlideName - Master Slide name
	 * @returns {ISlide} the new Slide
	 */
	addSlide(masterSlideName?: string): ISlide {
		let newSlide = new Slide({
			addSlide: this.addNewSlide,
			getSlide: this.getSlide,
			presLayout: this.presLayout,
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
	 * Adds a new slide master [layout] to the Presentation
	 * @param {ISlideMasterOptions} slideMasterOpts - layout definition
	 */
	defineSlideMaster(slideMasterOpts: ISlideMasterOptions) {
		if (!slideMasterOpts.title) throw Error('defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)')

		let newLayout: ISlideLayout = {
			presLayout: this.presLayout,
			name: slideMasterOpts.title,
			number: 1000 + this.slideLayouts.length + 1,
			slide: null,
			data: [],
			rels: [],
			relsChart: [],
			relsMedia: [],
			margin: slideMasterOpts.margin || DEF_SLIDE_MARGIN_IN,
			slideNumberObj: slideMasterOpts.slideNumber || null,
		}

		// STEP 1: Create the Slide Master/Layout
		genObj.createSlideObject(slideMasterOpts, newLayout)

		// STEP 2: Add it to layout defs
		this.slideLayouts.push(newLayout)

		// STEP 3: Add slideNumber to master slide (if any)
		if (newLayout.slideNumberObj && !this.masterSlide.slideNumberObj) this.masterSlide.slideNumberObj = newLayout.slideNumberObj
	}

	// HTML-TO-SLIDES METHODS

	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} tabEleId - HTMLElementID of the table
	 * @param {ITableToSlidesOpts} inOpts - array of options (e.g.: tabsize)
	 * @note `verbose` option is undocumented; used for verbose output of layout process
	 */
	tableToSlides(tableElementId: string, opts: ITableToSlidesOpts = {}) {
		genTable.genTableToSlides(
			this,
			tableElementId,
			opts,
			opts && opts.masterSlideName
				? this.slideLayouts.filter(layout => {
						return layout.name == opts.masterSlideName
				  })[0]
				: null
		)
	}
}
