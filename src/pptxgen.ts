/*\
|*|  :: pptxgen.ts ::
|*|
|*|  JavaScript framework that creates PowerPoint (pptx) presentations
|*|  https://github.com/gitbrent/PptxGenJS
|*|
|*|  This framework is released under the MIT Public License (MIT)
|*|
|*|  PptxGenJS (C) 2015-2020 Brent Ely -- https://github.com/gitbrent
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

import * as JSZip from 'jszip'
import Slide from './slide'
import {
	AlignH,
	AlignV,
	CHART_TYPE,
	ChartType,
	DEF_PRES_LAYOUT,
	DEF_PRES_LAYOUT_NAME,
	DEF_SLIDE_MARGIN_IN,
	EMU,
	JSZIP_OUTPUT_TYPE,
	OutputType,
	SCHEME_COLOR_NAMES,
	SHAPE_TYPE,
	SchemeColor,
	ShapeType,
	WRITE_OUTPUT_TYPE,
} from './core-enums'
import {
	IAddSlideOptions,
	ILayout,
	ILayoutProps,
	IPresentation,
	ISection,
	ISectionProps,
	ISlide,
	ISlideLayout,
	ISlideLib,
	ISlideMasterOptions,
	ISlideNumber,
	ITableToSlidesOpts,
} from './core-interfaces'
import * as genCharts from './gen-charts'
import * as genObj from './gen-objects'
import * as genMedia from './gen-media'
import * as genTable from './gen-tables'
import * as genXml from './gen-xml'

const VERSION = '3.3.0-beta-20200705:1153'

export default class PptxGenJS implements IPresentation {
	// Property getters/setters

	/**
	 * Presentation layout name
	 * Standard layouts:
	 * - 'LAYOUT_4x3'   (10"    x 7.5")
	 * - 'LAYOUT_16x9'  (10"    x 5.625")
	 * - 'LAYOUT_16x10' (10"    x 6.25")
	 * - 'LAYOUT_WIDE'  (13.33" x 7.5")
	 * Custom layouts:
	 * Use `pptx.defineLayout()` to create custom layouts (e.g.: 'A4')
	 * @type {string}
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 */
	private _layout: string
	public set layout(value: string) {
		let newLayout: ILayout = this.LAYOUTS[value]

		if (newLayout) {
			this._layout = value
			this._presLayout = newLayout
		} else {
			throw new Error('UNKNOWN-LAYOUT')
		}
	}
	public get layout(): string {
		return this._layout
	}

	/**
	 * PptxGenJS Library Version
	 */
	private _version: string = VERSION
	public get version(): string {
		return this._version
	}

	/**
	 * @type {string}
	 */
	private _author: string
	public set author(value: string) {
		this._author = value
	}
	public get author(): string {
		return this._author
	}

	/**
	 * @type {string}
	 */
	private _company: string
	public set company(value: string) {
		this._company = value
	}
	public get company(): string {
		return this._company
	}

	/**
	 * @type {string}
	 * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PPT will throw errors upon opening!)
	 */
	private _revision: string
	public set revision(value: string) {
		this._revision = value
	}
	public get revision(): string {
		return this._revision
	}

	/**
	 * @type {string}
	 */
	private _subject: string
	public set subject(value: string) {
		this._subject = value
	}
	public get subject(): string {
		return this._subject
	}

	/**
	 * @type {string}
	 */
	private _title: string
	public set title(value: string) {
		this._title = value
	}
	public get title(): string {
		return this._title
	}

	/**
	 * Whether Right-to-Left (RTL) mode is enabled
	 * @type {boolean}
	 */
	private _rtlMode: boolean
	public set rtlMode(value: boolean) {
		this._rtlMode = value
	}
	public get rtlMode(): boolean {
		return this._rtlMode
	}

	/** master slide layout object */
	private _masterSlide: ISlideLib
	public get masterSlide(): ISlideLib {
		return this._masterSlide
	}

	/** this Presentation's Slide objects */
	private _slides: ISlideLib[]
	public get slides(): ISlideLib[] {
		return this._slides
	}

	/** this Presentation's sections */
	private _sections: ISection[]
	public get sections(): ISection[] {
		return this._sections
	}

	/** slide layout definition objects, used for generating slide layout files */
	private _slideLayouts: ISlideLayout[]
	public get slideLayouts(): ISlideLayout[] {
		return this._slideLayouts
	}

	private LAYOUTS: object

	// Exposed class props
	private _alignH = AlignH
	public get AlignH(): typeof AlignH {
		return this._alignH
	}
	private _alignV = AlignV
	public get AlignV(): typeof AlignV {
		return this._alignV
	}
	private _chartType = ChartType
	public get ChartType(): typeof ChartType {
		return this._chartType
	}
	private _outputType = OutputType
	public get OutputType(): typeof OutputType {
		return this._outputType
	}
	private _presLayout: ILayout
	public get presLayout(): ILayout {
		return this._presLayout
	}
	private _schemeColor = SchemeColor
	public get SchemeColor(): typeof SchemeColor {
		return this._schemeColor
	}
	private _shapeType = ShapeType
	public get ShapeType(): typeof ShapeType {
		return this._shapeType
	}

	/**
	 * @depricated use `ChartType`
	 */
	private _charts = CHART_TYPE
	public get charts(): typeof CHART_TYPE {
		return this._charts
	}
	/**
	 * @depricated use `SchemeColor`
	 */
	private _colors = SCHEME_COLOR_NAMES
	public get colors(): typeof SCHEME_COLOR_NAMES {
		return this._colors
	}
	/**
	 * @depricated use `ShapeType`
	 */
	private _shapes = SHAPE_TYPE
	public get shapes(): typeof SHAPE_TYPE {
		return this._shapes
	}

	constructor() {
		// Set available layouts
		this.LAYOUTS = {
			LAYOUT_4x3: { name: 'screen4x3', width: 9144000, height: 6858000 } as ILayout,
			LAYOUT_16x9: { name: 'screen16x9', width: 9144000, height: 5143500 } as ILayout,
			LAYOUT_16x10: { name: 'screen16x10', width: 9144000, height: 5715000 } as ILayout,
			LAYOUT_WIDE: { name: 'custom', width: 12192000, height: 6858000 } as ILayout,
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
		//
		this._slideLayouts = [
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
		this._slides = []
		this._sections = []
		this._masterSlide = {
			addChart: null,
			addImage: null,
			addMedia: null,
			addNotes: null,
			addShape: null,
			addTable: null,
			addText: null,
			//
			presLayout: this._presLayout,
			id: null,
			rId: null,
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
	private addNewSlide = (masterName: string): ISlide => {
		// Continue using sections if the first slide using auto-paging has a Section
		let sectAlreadyInUse =
			this.sections.length > 0 && this.sections[this.sections.length - 1].slides.filter(slide => slide.number === this.slides[this.slides.length - 1].number).length > 0

		return this.addSlide({
			masterName: masterName,
			sectionTitle: sectAlreadyInUse ? this.sections[this.sections.length - 1].title : null,
		})
	}

	/**
	 * Provides an API for `addTableDefinition` to get slide reference by number
	 * @param {number} slideNum - slide number
	 * @return {ISlideLib} Slide
	 * @since 3.0.0
	 */
	private getSlide = (slideNum: number): ISlideLib => this.slides.filter(slide => slide.number === slideNum)[0]

	/**
	 * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
	 * @param {ISlideNumber} slideNum - slide number config
	 */
	private setSlideNumber = (slideNum: ISlideNumber) => {
		// 1: Add slideNumber to slideMaster1.xml
		this.masterSlide.slideNumberObj = slideNum

		// 2: Add slideNumber to DEF_PRES_LAYOUT_NAME layout
		this.slideLayouts.filter(layout => layout.name === DEF_PRES_LAYOUT_NAME)[0].slideNumberObj = slideNum
	}

	/**
	 * Create all chart and media rels for this Presentation
	 * @param {ISlideLib | ISlideLayout} slide - slide with rels
	 * @param {JSZIP} zip - JSZip instance
	 * @param {Promise<any>[]} chartPromises - promise array
	 */
	private createChartMediaRels = (slide: ISlideLib | ISlideLayout, zip: JSZip, chartPromises: Promise<any>[]) => {
		slide.relsChart.forEach(rel => chartPromises.push(genCharts.createExcelWorksheet(rel, zip)))
		slide.relsMedia.forEach(rel => {
			if (rel.type !== 'online' && rel.type !== 'hyperlink') {
				// A: Loop vars
				let data: string = rel.data && typeof rel.data === 'string' ? rel.data : ''

				// B: Users will undoubtedly pass various string formats, so correct prefixes as needed
				if (data.indexOf(',') === -1 && data.indexOf(';') === -1) data = 'image/png;base64,' + data
				else if (data.indexOf(',') === -1) data = 'image/png;base64,' + data
				else if (data.indexOf(';') === -1) data = 'image/png;' + data

				// C: Add media
				zip.file(rel.Target.replace('..', 'ppt'), data.split(',').pop(), { base64: true })
			}
		})
	}

	/**
	 * Create and export the .pptx file
	 * @param {string} exportName - output file type
	 * @param {Blob} blobContent - Blob content
	 * @return {Promise<string>} Promise with file name
	 */
	private writeFileToBrowser = (exportName: string, blobContent: Blob): Promise<string> => {
		// STEP 1: Create element
		let eleLink = document.createElement('a')
		eleLink.setAttribute('style', 'display:none;')
		eleLink.dataset.interception = 'off' // @see https://docs.microsoft.com/en-us/sharepoint/dev/spfx/hyperlinking
		document.body.appendChild(eleLink)

		// STEP 2: Download file to browser
		// DESIGN: Use `createObjectURL()` (or MS-specific func for IE11) to D/L files in client browsers (FYI: synchronously executed)
		if (window.navigator.msSaveOrOpenBlob) {
			// @see https://docs.microsoft.com/en-us/microsoft-edge/dev-guide/html5/file-api/blob
			let blob = new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
			eleLink.onclick = function () {
				window.navigator.msSaveOrOpenBlob(blob, exportName)
			}
			eleLink.click()

			// Clean-up
			document.body.removeChild(eleLink)

			// Done
			return Promise.resolve(exportName)
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
			return Promise.resolve(exportName)
		}
	}

	/**
	 * Create and export the .pptx file
	 * @param {WRITE_OUTPUT_TYPE} outputType - output file type
	 * @return {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} Promise with data or stream (node) or filename (browser)
	 */
	private exportPresentation = (outputType?: WRITE_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> => {
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

		// STEP 2: Wait for Promises (if any) then generate the PPTX file
		return Promise.all(arrMediaPromises).then(() => {
			// A: Add empty placeholder objects to slides that don't already have them
			this.slides.forEach(slide => {
				if (slide.slideLayout) genObj.addPlaceholdersToSlideLayouts(slide)
			})

			// B: Add all required folders and files
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
			zip.file('[Content_Types].xml', genXml.makeXmlContTypes(this.slides, this.slideLayouts, this.masterSlide)) // TODO: pass only `this` like below! 20200206
			zip.file('_rels/.rels', genXml.makeXmlRootRels())
			zip.file('docProps/app.xml', genXml.makeXmlApp(this.slides, this.company)) // TODO: pass only `this` like below! 20200206
			zip.file('docProps/core.xml', genXml.makeXmlCore(this.title, this.subject, this.author, this.revision)) // TODO: pass only `this` like below! 20200206
			zip.file('ppt/_rels/presentation.xml.rels', genXml.makeXmlPresentationRels(this.slides))
			zip.file('ppt/theme/theme1.xml', genXml.makeXmlTheme())
			zip.file('ppt/presentation.xml', genXml.makeXmlPresentation(this))
			zip.file('ppt/presProps.xml', genXml.makeXmlPresProps())
			zip.file('ppt/tableStyles.xml', genXml.makeXmlTableStyles())
			zip.file('ppt/viewProps.xml', genXml.makeXmlViewProps())

			// C: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
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

			// D: Create all Rels (images, media, chart data)
			this.slideLayouts.forEach(layout => {
				this.createChartMediaRels(layout, zip, arrChartPromises)
			})
			this.slides.forEach(slide => {
				this.createChartMediaRels(slide, zip, arrChartPromises)
			})
			this.createChartMediaRels(this.masterSlide, zip, arrChartPromises)

			// E: Wait for Promises (if any) then generate the PPTX file
			return Promise.all(arrChartPromises).then(() => {
				if (outputType === 'STREAM') {
					// A: stream file
					return zip.generateAsync({ type: 'nodebuffer' })
				} else if (outputType) {
					// B: Node [fs]: Output type user option or default
					return zip.generateAsync({ type: outputType })
				} else {
					// C: Browser: Output blob as app/ms-pptx
					return zip.generateAsync({ type: 'blob' })
				}
			})
		})
	}

	// EXPORT METHODS

	/**
	 * Export the current Presentation to stream
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
	 */
	stream(): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> {
		return this.exportPresentation('STREAM')
	}

	/**
	 * Export the current Presentation as JSZip content with the selected type
	 * @param {JSZIP_OUTPUT_TYPE} outputType - 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file content in selected type
	 */
	write(outputType: JSZIP_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> {
		return this.exportPresentation(outputType)
	}

	/**
	 * Export the current Presentation. Writes file to local file system if `fs` exists, otherwise, initiates download in browsers
	 * @param {string} exportName - file name
	 * @returns {Promise<string>} the presentation name
	 */
	writeFile(exportName?: string): Promise<string> {
		const fs = typeof require !== 'undefined' && typeof window === 'undefined' ? require('fs') : null // NodeJS
		let fileName = exportName ? (exportName.toString().toLowerCase().endsWith('.pptx') ? exportName : exportName + '.pptx') : 'Presentation.pptx'

		return this.exportPresentation(fs ? 'nodebuffer' : null).then(content => {
			if (fs) {
				// Node: Output
				return new Promise<string>((resolve, reject) => {
					fs.writeFile(fileName, content, err => {
						if (err) {
							reject(err)
						} else {
							resolve(fileName)
						}
					})
				})
			} else {
				// Browser: Output blob as app/ms-pptx
				return this.writeFileToBrowser(fileName, content as Blob)
			}
		})
	}

	// PRESENTATION METHODS

	/**
	 * Add a new Section to Presentation
	 * @param {ISectionProps} section - section properties
	 * @example pptx.addSection({ title:'Charts' });
	 */
	addSection(section: ISectionProps) {
		if (!section) console.warn('addSection requires an argument')
		else if (!section.title) console.warn('addSection requires a title')

		let newSection: ISection = {
			type: 'user',
			title: section.title,
			slides: [],
		}

		if (section.order) this.sections.splice(section.order, 0, newSection)
		else this._sections.push(newSection)
	}

	/**
	 * Add a new Slide to Presentation
	 * @param {IAddSlideOptions} options - slide options
	 * @returns {ISlide} the new Slide
	 */
	addSlide(options?: IAddSlideOptions): ISlide {
		// TODO: DEPRECATED: arg0 string "masterSlideName" dep as of 3.2.0
		let masterSlideName = typeof options === 'string' ? options : options && options.masterName ? options.masterName : ''

		let newSlide = new Slide({
			addSlide: this.addNewSlide,
			getSlide: this.getSlide,
			presLayout: this.presLayout,
			setSlideNum: this.setSlideNumber,
			slideId: this.slides.length + 256,
			slideRId: this.slides.length + 2,
			slideNumber: this.slides.length + 1,
			slideLayout: masterSlideName
				? this.slideLayouts.filter(layout => layout.name === masterSlideName)[0] || this.LAYOUTS[DEF_PRES_LAYOUT]
				: this.LAYOUTS[DEF_PRES_LAYOUT],
		})

		// A: Add slide to pres
		this._slides.push(newSlide)

		// B: Sections
		// B-1: Add slide to section (if any provided)
		if (options && options.sectionTitle) {
			let sect = this.sections.filter(section => section.title === options.sectionTitle)[0]
			if (!sect) console.warn(`addSlide: unable to find section with title: "${options.sectionTitle}"`)
			else sect.slides.push(newSlide)
		}
		// B-2: Handle slides without a section when sections are already is use ("loose" slides arent allowed, they all need a section)
		else if (this.sections && this.sections.length > 0 && (!options || !options.sectionTitle)) {
			let lastSect = this._sections[this.sections.length - 1]

			// CASE 1: The latest section is a default type - just add this one
			if (lastSect.type === 'default') lastSect.slides.push(newSlide)
			// CASE 2: There latest section is NOT a default type - create the defualt, add this slide
			else
				this._sections.push({
					type: 'default',
					title: `Default-${this.sections.filter(sect => sect.type === 'default').length + 1}`,
					slides: [newSlide],
				})
		}

		return newSlide
	}

	/**
	 * Create a custom Slide Layout in any size
	 * @param {ILayoutProps} layout - layout properties
	 * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
	 */
	defineLayout(layout: ILayoutProps) {
		// @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
		if (!layout) console.warn('defineLayout requires `{name, width, height}`')
		else if (!layout.name) console.warn('defineLayout requires `name`')
		else if (!layout.width) console.warn('defineLayout requires `width`')
		else if (!layout.height) console.warn('defineLayout requires `height`')
		else if (typeof layout.height !== 'number') console.warn('defineLayout `height` should be a number (inches)')
		else if (typeof layout.width !== 'number') console.warn('defineLayout `width` should be a number (inches)')

		this.LAYOUTS[layout.name] = { name: layout.name, width: Math.round(Number(layout.width) * EMU), height: Math.round(Number(layout.height) * EMU) }
	}

	/**
	 * Create a new slide master [layout] for the Presentation
	 * @param {ISlideMasterOptions} options - layout options
	 */
	defineSlideMaster(options: ISlideMasterOptions) {
		if (!options.title) throw Error('defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)')

		let newLayout: ISlideLayout = {
			presLayout: this.presLayout,
			name: options.title,
			number: 1000 + this.slideLayouts.length + 1,
			slide: null,
			data: [],
			rels: [],
			relsChart: [],
			relsMedia: [],
			margin: options.margin || DEF_SLIDE_MARGIN_IN,
			slideNumberObj: options.slideNumber || null,
		}

		// DEPRECATED:
		if (options.bkgd && !options.background) {
			options.background = {}
			if (typeof options.bkgd === 'string') options.background.fill = options.bkgd
			else {
				if (options.bkgd.data) options.background.data = options.bkgd.data
				if (options.bkgd.path) options.background.path = options.bkgd.path
				if (options.bkgd['src']) options.background.path = options.bkgd['src'] // @deprecated (drop in 4.x)
			}
			delete options.bkgd
		}

		// STEP 1: Create the Slide Master/Layout
		genObj.createSlideObject(options, newLayout)

		// STEP 2: Add it to layout defs
		this.slideLayouts.push(newLayout)

		// STEP 3: Add slideNumber to master slide (if any)
		if (newLayout.slideNumberObj && !this.masterSlide.slideNumberObj) this.masterSlide.slideNumberObj = newLayout.slideNumberObj
	}

	// HTML-TO-SLIDES METHODS

	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} eleId - table HTML element ID
	 * @param {ITableToSlidesOpts} options - generation options
	 */
	tableToSlides(eleId: string, options: ITableToSlidesOpts = {}) {
		// @note `verbose` option is undocumented; used for verbose output of layout process
		genTable.genTableToSlides(
			this,
			eleId,
			options,
			options && options.masterSlideName ? this.slideLayouts.filter(layout => layout.name === options.masterSlideName)[0] : null
		)
	}
}
