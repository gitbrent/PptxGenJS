/**
 *  :: pptxgen.ts ::
 *
 *  JavaScript framework that creates PowerPoint (pptx) presentations
 *  https://github.com/gitbrent/PptxGenJS
 *
 *  This framework is released under the MIT Public License (MIT)
 *
 *  PptxGenJS (C) 2015-present Brent Ely -- https://github.com/gitbrent
 *
 *  Some code derived from the OfficeGen project:
 *  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in all
 *  copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 *  SOFTWARE.
 */

/**
 * Units of Measure used in PowerPoint documents
 *
 * PowerPoint units are in `DXA` (except for font sizing)
 * - 1 inch is 1440 DXA
 * - 1 inch is 72 points
 * - 1 DXA is 1/20th's of a point
 * - 20 DXA is 1 point
 *
 * Another form of measurement using is an `EMU`
 * - 914400 EMUs is 1 inch
 * 12700 EMUs is 1 point
 *
 * @see https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
 */

/**
 * Object Layouts
 *
 * - 16x9 (10" x 5.625")
 * - 16x10 (10" x 6.25")
 * - 4x3 (10" x 7.5")
 * - Wide (13.33" x 7.5")
 * - [custom] (any size)
 *
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh273476(v=office.14)
 */

import JSZip from 'jszip'
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
	OutputType,
	SCHEME_COLOR_NAMES,
	SHAPE_TYPE,
	SchemeColor,
	ShapeType,
	WRITE_OUTPUT_TYPE,
} from './core-enums'
import {
	AddSlideProps,
	IPresentationProps,
	PresLayout,
	PresSlide,
	SectionProps,
	SlideLayout,
	SlideMasterProps,
	SlideNumberProps,
	TableToSlidesProps,
	ThemeProps,
	WriteBaseProps,
	WriteFileProps,
	WriteProps,
} from './core-interfaces'
import * as genCharts from './gen-charts'
import * as genObj from './gen-objects'
import * as genMedia from './gen-media'
import * as genTable from './gen-tables'
import * as genXml from './gen-xml'

const VERSION = '3.13.0-beta.0-20230416-2140'

export default class PptxGenJS implements IPresentationProps {
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
	public set layout (value: string) {
		const newLayout: PresLayout = this.LAYOUTS[value]

		if (newLayout) {
			this._layout = value
			this._presLayout = newLayout
		} else {
			throw new Error('UNKNOWN-LAYOUT')
		}
	}

	public get layout (): string {
		return this._layout
	}

	/**
	 * PptxGenJS Library Version
	 */
	private readonly _version: string = VERSION
	public get version (): string {
		return this._version
	}

	/**
	 * @type {string}
	 */
	private _author: string
	public set author (value: string) {
		this._author = value
	}

	public get author (): string {
		return this._author
	}

	/**
	 * @type {string}
	 */
	private _company: string
	public set company (value: string) {
		this._company = value
	}

	public get company (): string {
		return this._company
	}

	/**
	 * @type {string}
	 * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PPT will throw errors upon opening!)
	 */
	private _revision: string
	public set revision (value: string) {
		this._revision = value
	}

	public get revision (): string {
		return this._revision
	}

	/**
	 * @type {string}
	 */
	private _subject: string
	public set subject (value: string) {
		this._subject = value
	}

	public get subject (): string {
		return this._subject
	}

	/**
	 * @type {ThemeProps}
	 */
	private _theme: ThemeProps
	public set theme (value: ThemeProps) {
		this._theme = value
	}

	public get theme (): ThemeProps {
		return this._theme
	}

	/**
	 * @type {string}
	 */
	private _title: string
	public set title (value: string) {
		this._title = value
	}

	public get title (): string {
		return this._title
	}

	/**
	 * Whether Right-to-Left (RTL) mode is enabled
	 * @type {boolean}
	 */
	private _rtlMode: boolean
	public set rtlMode (value: boolean) {
		this._rtlMode = value
	}

	public get rtlMode (): boolean {
		return this._rtlMode
	}

	/** master slide layout object */
	private readonly _masterSlide: PresSlide
	public get masterSlide (): PresSlide {
		return this._masterSlide
	}

	/** this Presentation's Slide objects */
	private readonly _slides: PresSlide[]
	public get slides (): PresSlide[] {
		return this._slides
	}

	/** this Presentation's sections */
	private readonly _sections: SectionProps[]
	public get sections (): SectionProps[] {
		return this._sections
	}

	/** slide layout definition objects, used for generating slide layout files */
	private readonly _slideLayouts: SlideLayout[]
	public get slideLayouts (): SlideLayout[] {
		return this._slideLayouts
	}

	private LAYOUTS: { [key: string]: PresLayout }

	// Exposed class props
	private readonly _alignH = AlignH
	public get AlignH (): typeof AlignH {
		return this._alignH
	}

	private readonly _alignV = AlignV
	public get AlignV (): typeof AlignV {
		return this._alignV
	}

	private readonly _chartType = ChartType
	public get ChartType (): typeof ChartType {
		return this._chartType
	}

	private readonly _outputType = OutputType
	public get OutputType (): typeof OutputType {
		return this._outputType
	}

	private _presLayout: PresLayout
	public get presLayout (): PresLayout {
		return this._presLayout
	}

	private readonly _schemeColor = SchemeColor
	public get SchemeColor (): typeof SchemeColor {
		return this._schemeColor
	}

	private readonly _shapeType = ShapeType
	public get ShapeType (): typeof ShapeType {
		return this._shapeType
	}

	/**
	 * @depricated use `ChartType`
	 */
	private readonly _charts = CHART_TYPE
	public get charts (): typeof CHART_TYPE {
		return this._charts
	}

	/**
	 * @depricated use `SchemeColor`
	 */
	private readonly _colors = SCHEME_COLOR_NAMES
	public get colors (): typeof SCHEME_COLOR_NAMES {
		return this._colors
	}

	/**
	 * @depricated use `ShapeType`
	 */
	private readonly _shapes = SHAPE_TYPE
	public get shapes (): typeof SHAPE_TYPE {
		return this._shapes
	}

	constructor () {
		const layout4x3: PresLayout = { name: 'screen4x3', width: 9144000, height: 6858000 }
		const layout16x9: PresLayout = { name: 'screen16x9', width: 9144000, height: 5143500 }
		const layout16x10: PresLayout = { name: 'screen16x10', width: 9144000, height: 5715000 }
		const layoutWide: PresLayout = { name: 'custom', width: 12192000, height: 6858000 }
		// Set available layouts
		this.LAYOUTS = {
			LAYOUT_4x3: layout4x3,
			LAYOUT_16x9: layout16x9,
			LAYOUT_16x10: layout16x10,
			LAYOUT_WIDE: layoutWide,
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
			_sizeW: this.LAYOUTS[DEF_PRES_LAYOUT].width,
			_sizeH: this.LAYOUTS[DEF_PRES_LAYOUT].height,
			width: this.LAYOUTS[DEF_PRES_LAYOUT].width,
			height: this.LAYOUTS[DEF_PRES_LAYOUT].height,
		}
		this._rtlMode = false
		//
		this._slideLayouts = [
			{
				_margin: DEF_SLIDE_MARGIN_IN,
				_name: DEF_PRES_LAYOUT_NAME,
				_presLayout: this._presLayout,
				_rels: [],
				_relsChart: [],
				_relsMedia: [],
				_slide: null,
				_slideNum: 1000,
				_slideNumberProps: null,
				_slideObjects: [],
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
			_name: null,
			_presLayout: this._presLayout,
			_rId: null,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slideId: null,
			_slideLayout: null,
			_slideNum: null,
			_slideNumberProps: null,
			_slideObjects: [],
		}
	}

	/**
	 * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
	 * @param {AddSlideProps} options - slide masterName and/or sectionTitle
	 * @return {PresSlide} new Slide
	 */
	private readonly addNewSlide = (options?: AddSlideProps): PresSlide => {
		// Continue using sections if the first slide using auto-paging has a Section
		const sectAlreadyInUse =
			this.sections.length > 0 &&
			this.sections[this.sections.length - 1]._slides.filter(slide => slide._slideNum === this.slides[this.slides.length - 1]._slideNum).length > 0

		options.sectionTitle = sectAlreadyInUse ? this.sections[this.sections.length - 1].title : null

		return this.addSlide(options)
	}

	/**
	 * Provides an API for `addTableDefinition` to get slide reference by number
	 * @param {number} slideNum - slide number
	 * @return {PresSlide} Slide
	 * @since 3.0.0
	 */
	private readonly getSlide = (slideNum: number): PresSlide => this.slides.filter(slide => slide._slideNum === slideNum)[0]

	/**
	 * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
	 * @param {SlideNumberProps} slideNum - slide number config
	 */
	private readonly setSlideNumber = (slideNum: SlideNumberProps): void => {
		// 1: Add slideNumber to slideMaster1.xml
		this.masterSlide._slideNumberProps = slideNum

		// 2: Add slideNumber to DEF_PRES_LAYOUT_NAME layout
		this.slideLayouts.filter(layout => layout._name === DEF_PRES_LAYOUT_NAME)[0]._slideNumberProps = slideNum
	}

	/**
	 * Create all chart and media rels for this Presentation
	 * @param {PresSlide | SlideLayout} slide - slide with rels
	 * @param {JSZip} zip - JSZip instance
	 * @param {Promise<string>[]} chartPromises - promise array
	 */
	private readonly createChartMediaRels = (slide: PresSlide | SlideLayout, zip: JSZip, chartPromises: Array<Promise<string>>): void => {
		slide._relsChart.forEach(rel => chartPromises.push(genCharts.createExcelWorksheet(rel, zip)))
		slide._relsMedia.forEach(rel => {
			if (rel.type !== 'online' && rel.type !== 'hyperlink') {
				// A: Loop vars
				let data: string = rel.data && typeof rel.data === 'string' ? rel.data : ''

				// B: Users will undoubtedly pass various string formats, so correct prefixes as needed
				if (!data.includes(',') && !data.includes(';')) data = 'image/png;base64,' + data
				else if (!data.includes(',')) data = 'image/png;base64,' + data
				else if (!data.includes(';')) data = 'image/png;' + data

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
	private readonly writeFileToBrowser = async (exportName: string, blobContent: Blob): Promise<string> => {
		// STEP 1: Create element
		const eleLink = document.createElement('a')
		eleLink.setAttribute('style', 'display:none;')
		eleLink.dataset.interception = 'off' // @see https://docs.microsoft.com/en-us/sharepoint/dev/spfx/hyperlinking
		document.body.appendChild(eleLink)

		// STEP 2: Download file to browser
		// DESIGN: Use `createObjectURL()` to D/L files in client browsers (FYI: synchronously executed)
		if (window.URL.createObjectURL) {
			const url = window.URL.createObjectURL(new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }))
			eleLink.href = url
			eleLink.download = exportName
			eleLink.click()

			// Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
			setTimeout(() => {
				window.URL.revokeObjectURL(url)
				document.body.removeChild(eleLink)
			}, 100)

			// Done
			return await Promise.resolve(exportName)
		}
	}

	/**
	 * Create and export the .pptx file
	 * @param {WRITE_OUTPUT_TYPE} outputType - output file type
	 * @return {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} Promise with data or stream (node) or filename (browser)
	 */
	private readonly exportPresentation = async (props: WriteProps): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> => {
		const arrChartPromises: Array<Promise<string>> = []
		let arrMediaPromises: Array<Promise<string>> = []
		const zip = new JSZip()

		// STEP 1: Read/Encode all Media before zip as base64 content, etc. is required
		this.slides.forEach(slide => {
			arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(slide))
		})
		this.slideLayouts.forEach(layout => {
			arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(layout))
		})
		arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(this.masterSlide))

		// STEP 2: Wait for Promises (if any) then generate the PPTX file
		return await Promise.all(arrMediaPromises).then(async () => {
			// A: Add empty placeholder objects to slides that don't already have them
			this.slides.forEach(slide => {
				if (slide._slideLayout) genObj.addPlaceholdersToSlideLayouts(slide)
			})

			// B: Add all required folders and files
			zip.folder('_rels')
			zip.folder('docProps')
			zip.folder('ppt').folder('_rels')
			zip.folder('ppt/charts').folder('_rels')
			zip.folder('ppt/embeddings')
			// zip.folder('ppt/media')
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
			zip.file('ppt/charts/style1.xml', getStyle1())
			zip.file('ppt/charts/colors1.xml', getColors1())
			zip.file('ppt/_rels/presentation.xml.rels', genXml.makeXmlPresentationRels(this.slides))
			zip.file('ppt/theme/theme1.xml', genXml.makeXmlTheme(this))
			zip.file('ppt/theme/theme2.xml', Theme2())
			zip.file('ppt/presentation.xml', genXml.makeXmlPresentation(this))
			zip.file('ppt/presProps.xml', genXml.makeXmlPresProps())
			zip.file('ppt/tableStyles.xml', genXml.makeXmlTableStyles())
			zip.file('ppt/viewProps.xml', genXml.makeXmlViewProps())
			// zip.file('ppt/media/image1.png', 'iVBORw0KGgoAAAANSUhEUgAABcEAAALKCAYAAAD6TZk0AAAAAXNSR0IArs4c6QAAAIRlWElmTU0AKgAAAAgABQESAAMAAAABAAEAAAEaAAUAAAABAAAASgEbAAUAAAABAAAAUgEoAAMAAAABAAIAAIdpAAQAAAABAAAAWgAAAAAAAABgAAAAAQAAAGAAAAABAAOgAQADAAAAAQABAACgAgAEAAAAAQAABcGgAwAEAAAAAQAAAsoAAAAANXirAAAAAAlwSFlzAAAOxAAADsQBlSsOGwAAQABJREFUeAHs3Qu0ZXV9H/A7PIZhEHnLo1AEYqA8zJrwECgQjWlCEEkCSVMMCUsCwYghvlokCIoQopaISTCF1SWpC5dNE0uwgSAtBBKFEBAkExGHDATF8hCUIswIDMz0+7uzNx6O9869N+5z77nnfv5rffn/93/vs8/en33XrObX7f+MjS2QtmLFinUL5FbdJgECBAgQIECAAAECBAgQIECAAAECBAg0AhuRIECAAAECBAgQIECAAAECBAgQIECAAAECoyqgCD6qT9Z9ESBAgAABAgQIECBAgAABAgQIECBAgMCYIrg/AgIECBAgQIAAAQIECBAgQIAAAQIECBAYWQFF8JF9tG6MAAECBAgQIECAAAECBAgQIECAAAECBBTB/Q0QIECAAAECBAgQIECAAAECBAgQIECAwMgKKIKP7KN1YwQIECBAgAABAgQIECBAgAABAgQIECCgCO5vgAABAgQIECBAgAABAgQIECBAgAABAgRGVkARfGQfrRsjQIAAAQIECBAgQIAAAQIECBAgQIAAAUVwfwMECBAgQIAAAQIECBAgQIAAAQIECBAgMLICiuAj+2jdGAECBAgQIECAAAECBAgQIECAAAECBAgogvsbIECAAAECBAgQIECAAAECBAgQIECAAIGRFVAEH9lH68YIECBAgAABAgQIECBAgAABAgQIECBAQBHc3wABAgQIECBAgAABAgQIECBAgAABAgQIjKzAJh3fWZ3vgGS/5Jnk75LHkv62OBPLkr2Te5LlyZqkbVWcPzTZKbkjeSjpb2dlYlVyaf8O2wQIECBAgAABAgQIECBAgAABAgQIECBAoGuB1+SEdyXrelJF6l9Jetsh2fh20nvcymzv3nPQ5zJem6xOnkuOSXrbqdmoz7+td3JD4xUrVtTxGgECBAgQIECAAAECBAgQIECAAAECBAgQmLHAonzi9qQKzRck+ycnJ20Re5eMq+2QfDd5NjkjOTD5WFKfqwJ6tSOS2j49WZpUgfyWpG17ZfB0cm07MZ1eEXw6So4hQIAAAQIECBAgQIAAAQIECBAgQIAAgYkEfiyTVbi+oW/nRc38e5r5M5vt9/Udd1szv2/6Kn7XuV6dVPtU8sj4aGxs4/S3Jo8nOzZz0+oUwafF5CACBAgQIECAAAECBAgQIECAAAECBAiMlEBXP4xZb2dX+9L67qX/tm9r/3gzc2zTX/XSEesH1zTbR6V/oRkvafp6y/z5Znx2+sOS05KJ1hpvDtMRIECAAAECBAgQIECAAAECBAgQIECAAIGxsa6K4P+vwdyjD7V9g7t9a3vn7H8xua/vuPub7e3T392Ma7mUKngfndSb4gcl5yVXJFcnGgECBAgQIECAAAECBAgQIECAAAECBAgQmBWBbfItTyVrkt9Kak3wn0nqje9a2uRvkmr19vaT46OX/+e4bNZxFzfTl6SvYnnNLU/2S+5Nqli+ZXJwckKyZzKtZjmUaTE5iAABAgQIECBAgAABAgQIECBAgAABAgQmEfj3ma8fvazCdZt2+7rmM4+mX9WMe7s3ZaM+85Geyc0zbt8gvzTjWibl8OSTSR1bP7pZc29JpmyK4FMSOYAAAQIECBAgQIAAAQIECBAgQIAAAQIEphCoN8J/IvmlpN7WPiSpgvXlSbWvJLW9aW30tJMyrvn2BzR7do0vh1L7Lkz2SWp8TlLnuDNZkUzZFMGnJHIAAQIECBAgQIAAAQIECBAgQIAAAQIERk6gqzXBW5ha6qSWPvnz5I7kjUm1WtO7Wvtjlges33zpv+32gy/NrB9sl67WAK9i9/lJLYtS7fqkll65Oal1yBclGgECBAgQIECAAAECBAgQIECAAAECBAgQeJlA10Xw3pP/SDbemdSSKJ9tdlzb9PXmd9uWZnBi8kzSLpvS7qs3yLdO6vgqetfyJ9WWrO/Gi981X2+HawQIECBAgAABAgQIECBAgAABAgQIECBAYGAC9Xb2B5K3Jr+fPJ5UcfpXk7btkkG9LV7zlyVnJrc32xel720nZ6OOO6Nnst76rh/MvDI5MHkguTGZslkOZUoiBxAgQIAAAQIECBAgQIAAAQIECBAgQIDABgTuyr4qWre5N+NfmOD4ZZmrtcHXJnVs/VjmBUlvq2L5U0n/m+F1zLlJ+/b3yoz3r8mpmiL4VEL2EyBAgAABAgQIECBAgAABAgQIECBAgMBUAjvmgB9NXjHVgdm/VbLrJMfVvuOT+qHNidpmmdx5oh2TzSmCTyZjngABAgQIECBAgAABAgQIECBAgAABAgTmvYAi+Lx/hG6AAAECBAgQIECAAAECBAgQIECAAAECMxYY5A9jzvhifIAAAQIECBAgQIAAAQIECBAgQIAAAQIECHQpoAjepaZzESBAgAABAgQIECBAgAABAgQIECBAgMBQCSiCD9XjcDEECBAgQIAAAQIECBAgQIAAAQIECBAg0KWAIniXms5FgAABAgQIECBAgAABAgQIECBAgAABAkMloAg+VI/DxRAgQIAAAQIECBAgQIAAAQIECBAgQIBAlwKK4F1qOhcBAgQIECBAgAABAgQIECBAgAABAgQIDJWAIvhQPQ4XQ4AAAQIECBAgQIAAAQIECBAgQIAAAQJdCiiCd6npXAQIECBAgAABAgQIECBAgAABAgQIECAwVAKK4EP1OFwMAQIECBAgQIAAAQIECBAgQIAAAQIECHQpoAjepaZzESBAgAABAgQIECBAgAABAgQIECBAgMBQCSiCD9XjcDEECBAgQIAAAQIECBAgQIAAAQIECBAg0KWAIniXms5FgAABAgQIECBAgAABAgQIECBAgAABAkMloAg+VI/DxRAgQIAAAQIECBAgQIAAAQIECBAgQIBAlwKK4F1qOhcBAgQIECBAgAABAgQIECBAgAABAgQIDJWAIvhQPQ4XQ4AAAQIECBAgQIAAAQIECBAgQIAAAQJdCiiCd6npXAQIECBAgAABAgQIECBAgAABAgQIECAwVAKK4EP1OFwMAQIECBAgQIAAAQIECBAgQIAAAQIECHQpoAjepaZzESBAgAABAgQIECBAgAABAgQIECBAgMBQCSiCD9XjcDEECBAgQIAAAQIECBAgQIAAAQIECBAg0KWAIniXms5FgAABAgQIECBAgAABAgQIECBAgAABAkMloAg+VI/DxRAgQIAAAQIECBAgQIAAAQIECBAgQIBAlwKK4F1qOhcBAgQIECBAgAABAgQIECBAgAABAgQIDJWAIvhQPQ4XQ4AAAQIECBAgQIAAAQIECBAgQIAAAQJdCiiCd6npXAQIECBAgAABAgQIECBAgAABAgQIECAwVAKK4EP1OFwMAQIECBAgQIAAAQIECBAgQIAAAQIECHQpoAjepaZzESBAgAABAgQIECBAgAABAgQIECBAgMBQCSiCD9XjcDEECBAgQIAAAQIECBAgQIAAAQIECBAg0KWAIniXms5FgAABAgQIECBAgAABAgQIECBAgAABAkMloAg+VI/DxRAgQIAAAQIECBAgQIAAAQIECBAgQIBAlwKbdHmynGtRclDymua8/5D+nmbc2y3OxrJk76T2L0/WJG2r4vyhyU7JHclDSX87KxOrkkv7d9gmQIAAAQIECBAgQIAAAQIECBAgQIAAAQJdC/yrnPDLybq+/Pds975xfki2v913zMps75607XMZrE1WJ88lxyS97dRs1Pe8rXdyQ+MVK1bU8RoBAgQIECBAgAABAgQIECBAgAABAgQIEPgXCfxZPlWF5suTesv74OTOpOb+Q1Jth+S7ybPJGcmByceSOuaupNoRSW2fnixNqkB+S9K2vTJ4Orm2nZhOrwg+HSXHECBAgAABAgQIECBAgAABAgQIECBAgMBkAt/Mjipub9FzwMkZV0H7/GbuzGb7fc12293WzO+bvorf9ZlXJ9U+lTwyPhob2zj9rcnjyY7N3LQ6RfBpMTmIAAECBAgQIECAAAECBAgQIECAAAECIyXQu0zJD3tjT+QEtdb3/j0n+plm/NWmP7bpr+o5pobXNNtHpX+hGS9p+lpn/PlmfHb6w5LTkseaOR0BAgQIECBAgAABAgQIECBAgAABAgQIEBi4wC/lG2od71ru5O3Jf0vqje4bknqDu9o/Jm2Re3yi+c+J6evY9ye1REqN/yipgve3kv+R1A9uVjH8k8mMmzfBZ0zmAwQIECBAgAABAgQIECBAgAABAgQIECDQJ/CL2a4CdpsvZdy+0V2H1tvbT9agrx2X7frMxc38JelfbOaWp98vuTe5P9kyqfXGT0j2TKbVFMGnxeQgAgQIECBAgAABAgQIECBAgAABAgQIEJhEoArSdydVzL4uqXW7a/z55JVJtUeTVeOjl//nTdmsYz/SM715xu2635dmXG+QH57Um+B17Oqk5t6STNkUwackcgABAgQIECBAgAABAgQIECBAgAABAgQITCJQb3uvTGq5kjc3x2ydvtb6roL1Fc3cV5rtTZvttjupmX9PO9HTH93suzD9Ps34nPR1jjuTFcmUTRF8SiIHECBAgAABAgQIECBAgAABAgQIECBAgMAkAm/MfG+xuz3sFRmsSR5pJm5MX8f9eLPddvUGeM3XEie9bbtsPJzUsipV9K79ddxBSbXfT6rwXj+eucGmCL5BHjsJECBAgAABAgQIECBAgAABAgQIECAwkgIbdXRX7Zvd+/adb7dsb5I81cxf2/T15nfblmZQP4z5TFLLqPS2y7NRb5TX8VVMfyGpVm+eV6vid81XYVwjQIAAAQIECBAgQIAAAQIECBAgQIAAAQIDEahCdbsG+FUZ/1bywaR+zLIK1Ocn1XZJ6ocxa+6y5Mzk9mb7ovS97eRs1HFn9EzukfGLyZXJgckDSb1dPmXzJviURA4gQIAAAQIECBAgQIAAAQIECBAgQIAAgQ0IHJJ9tyZVuG7z/zL+cLJx0rZlGdTa4GuTOu7R5IKkt1WxvN4e738zvI45N2nf/l6Z8f41OVVTBJ9KyH4CBAgQIECAAAECBAgQIECAAAECBAgQmI5ArQP+I8lOyYbW6t4q+3dNJmq17/hkm4l2Zm6zZOdJ9k04rQg+IYtJAgQIECBAgAABAgQIECBAgAABAgQIEBgFAUXwUXiK7oEAAQIECBAgQIAAAQIECBAgQIAAAQIzE+jqhzFn9q2OJkCAAAECBAgQIECAAAECBAgQIECAAAECsyCgCD4LyL6CAAECBAgQIECAAAECBAgQIECAAAECBOZGQBF8btx9KwECBAgQIECAAAECBAgQIECAAAECBAjMgoAi+Cwg+woCBAgQIECAAAECBAgQIECAAAECBAgQmBsBRfC5cfetBAgQIECAAAECBAgQIECAAAECBAgQIDALAorgs4DsKwgQIECAAAECBAgQIECAAAECBAgQIEBgbgQUwefG3bcSIECAAAECBAgQIECAAAECBAgQIECAwCwIKILPArKvIECAAAECBAgQIECAAAECBAgQIECAAIG5EVAEnxt330qAAAECBAgQIECAAAECBAgQIECAAAECsyCgCD4LyL6CAAECBAgQIECAAAECBAgQIECAAAECBOZGQBF8btx9KwECBAgQIECAAAECBAgQIECAAAECBAjMgoAi+Cwg+woCBAgQIECAAAECBAgQIECAAAECBAgQmBsBRfC5cfetBAgQIECAAAECBAgQIECAAAECBAgQIDALAorgs4DsKwgQIECAAAECBAgQIECAAAECBAgQIEBgbgQUwefG3bcSIECAAAECBAgQIECAAAECBAgQIECAwCwIKILPArKvIECAAAECBAgQIECAAAECBAgQIECAAIG5EVAEnxt330qAAAECBAgQIECAAAECBAgQIECAAAECsyCgCD4LyL6CAAECBAgQIECAAAECBAgQIECAAAECBOZGQBF8btx9KwECBAgQIECAAAECBAgQIECAAAECBAjMgoAi+Cwg+woCBAgQIECAAAECBAgQIECAAAECBAgQmBsBRfC5cfetBAgQIECAAAECBAgQIECAAAECBAgQIDALAorgs4DsKwgQIECAAAECBAgQIECAAAECBAgQIEBgbgQUwefG3bcSIECAAAECBAgQIECAAAECBAgQIECAwCwIKILPArKvIECAAAECBAgQIECAAAECBAgQIECAAIG5EVAEnxt330qAAAECBAgQIECAAAECBAgQIECAAAECsyCgCD4LyL6CAAECBAgQIECAAAECBAgQIECAAAECBOZGQBF8btx9KwECBAgQIECAAAECBAgQIECAAAECBAjMgoAi+Cwg+woCBAgQIECAAAECBAgQIECAAAECBAgQmBsBRfC5cfetBAgQIECAAAECBAgQIECAAAECBAgQIDALApt09B0b5zwHJ4smOd+Tmf9az77FGS9L9k7uSZYna5K2VXH+0GSn5I7koaS/nZWJVcml/TtsEyBAgAABAgQIECBAgAABAgQIECBAgACBLgX2z8nWbSBV6G7bIRl8O+k9fmW2d28PSP+5ZG2yOnkuOSbpbadmoz7/tt7JDY1XrFhRx2sECBAgQIAAAQIECBAgQIAAAQIECBAgQGDGApvlE7+enNaX/5rtKj5fl1TbIflu8mxyRnJg8rGkjrkrqXZEUtunJ0uTKpDfkrRtrwyeTq5tJ6bTK4JPR8kxBAgQIECAAAECBAgQIECAAAECBAgQIDATgT/PwVXQ/rnmQ2c22+9rttvutmZ+3/RV/K7PvDqp9qnkkfHR2Fgtu3Jr8niyYzM3rU4RfFpMDiJAgAABAgQIECBAgAABAgQIECBAgMBICQzyhzH3iNQvJA8kf9moHdv0VzV9213TDI5K/0IzXtL0tc7488347PSHJfXG+WPNnI4AAQIECBAgQIAAAQIECBAgQIAAAQIECEwoMMgieL31XW9u1w9X1vre1XZOXkzuq42edn8z3j793c24lkupgvfRSb0pflByXnJFcnWiESBAgAABAgQIECBAgAABAgQIECBAgACBORHYKt9aa3/X2t01blu9vf1ku9HTH5dxLYFycTN3Sfoqltfc8mS/5N6kiuVbJgcnJyR7JtNqlkOZFpODCBAgQIAAAQIECBAgQIAAAQIECBAgQGAaAu/NMVXA/qO+Yx/N9qq+udp8U1LHf6Q2mrZ5+nbd73qbvJZJOTz5ZFLHrk5q7i3JlE0RfEoiBxAgQIAAAQIECBAgQIAAAQIECBAgQIDANAQ2yTHfSGoJlNf0Hf+VbFcBe9O++ZOa+ff0zddmLYdSn7kw2acZn5O+znFnsiKZsimCT0nkAAIECBAgQIAAAQIECBAgQIAAAQIECIycwCDWBP/FKO2WXJf8U59Y+2OWB/TNt9sP9s1vl+1aA7yK3ecntSxKteuTNcnNSf0AZ/14pkaAAAECBAgQIECAAAECBAgQIECAAAECBF4mMIgi+Lubb/iDl33T+o1rm7l687ttSzM4MXkmqcJ5b7s8G1sndXwVvWv5k2pL1nfjxe+arzfFNQIECBAgQIAAAQIECBAgQIAAAQIECBAgMFCBI3P2Kkh/dZJv2SXz9cOYdcxlyZnJ7c32Rel728nZqOPO6Jmst75fTK5MDkweSG5MpmyWQ5mSyAEECBAgQIAAAQIECBAgQIAAAQIECBAgMIXAJ7O/Cte/voHjlmVfrQ1ea4bXsfVjmRckva2K5U8l/W+G1zHnJu3b3ysz3r8mp2qK4FMJ2U+AAAECBAgQIECAAAECBAgQIECAAAECUwnUj2JuMdVBzf6t0u86ybG17/hkm0n2b5b5nSfZN+G0IviELCYJECBAgAABAgQIECBAgAABAgQIECBAYBQEFMFH4Sm6BwIECBAgQIAAAQIECBAgQIAAAQIECMxMYBA/jDmzK3A0AQIECBAgQIAAAQIECBAgQIAAAQIECBAYkIAi+IBgnZYAAQIECBAgQIAAAQIECBAgQIAAAQIE5l5AEXzun4ErIECAAAECBAgQIECAAAECBAgQIECAAIEBCSiCDwjWaQkQIECAAAECBAgQIECAAAECBAgQIEBg7gUUwef+GbgCAgQIECBAgAABAgQIECBAgAABAgQIEBiQgCL4gGCdlgABAgQIECBAgAABAgQIECBAgAABAgTmXkARfO6fgSsgQIAAAQIECBAgQIAAAQIECBAgQIAAgQEJKIIPCNZpCRAgQIAAAQIECBAgQIAAAQIECBAgQGDuBRTB5/4ZuAICBAgQIECAAAECBAgQIECAAAECBAgQGJCAIviAYJ2WAAECBAgQIECAAAECBAgQIECAAAECBOZeQBF87p+BKyBAgAABAgQIECBAgAABAgQIECBAgACBAQkogg8I1mkJECBAgAABAgQIECBAgAABAgQIECBAYO4FFMHn/hm4AgIECBAgQIAAAQIECBAgQIAAAQIECBAYkIAi+IBgnZYAAQIECBAgQIAAAQIECBAgQIAAAQIE5l5AEXzun4ErIECAAAECBAgQIECAAAECBAgQIECAAIEBCSiCDwjWaQkQIECAAAECBAgQIECAAAECBAgQIEBg7gUUwef+GbgCAgQIECBAgAABAgQIECBAgAABAgQIEBiQgCL4gGCdlgABAgQIECBAgAABAgQIECBAgAABAgTmXkARfO6fgSsgQIAAAQIECBAgQIAAAQIECBAgQIAAgQEJKIIPCNZpCRAgQIAAAQIECBAgQIAAAQIECBAgQGDuBRTB5/4ZuAICBAgQIECAAAECBAgQIECAAAECBAgQGJCAIviAYJ2WAAECBAgQIECAAAECBAgQIECAAAECBOZeQBF87p+BKyBAgAABAgQIECBAgAABAgQIECBAgACBAQkogg8I1mkJECBAgAABAgQIECBAgAABAgQIECBAYO4FFMHn/hm4AgIECBAgQIAAAQIECBAgQIAAAQIECBAYkIAi+IBgnZYAAQIECBAgQIAAAQIECBAgQIAAAQIE5l5AEXzun4ErIECAAAECBAgQIECAAAECBAgQIECAAIEBCSiCDwjWaQkQIECAAAECBAgQIECAAAECBAgQIEBg7gUUwef+GbgCAgQIECBAgAABAgQIECBAgAABAgQIEBiQgCL4gGCdlgABAgQIECBAgAABAgQIECBAgAABAgTmXmCTAVzCopxz/2SfZEnyteSOpLctzsayZO/knmR5siZpWxXnD012SuqzDyX97axMrEou7d9hmwABAgQIECBAgAABAgQIECBAgAABAgQIDEKgitq3Juv68is9X3ZIxt/u278y27v3HPO5jNcmq5PnkmOS3nZqNuo73tY7uaHxihUr6niNAAECBAgQIECAAAECBAgQIECAAAECBAj8iwS2zqfuT6rY/J+T1yX7Jb+RvDaptkPy3eTZ5IzkwORjSX3mrqTaEUltn54sTapAfkvStr0yeDq5tp2YTq8IPh0lxxAgQIAAAQIECBAgQIAAAQIECBAgQIDAZAK1PEkVrz842QGZPzOpY97Xd8xtzfy+6av4Xce8Oqn2qeSR8dHY2Mbp603zx5Mdm7lpdYrg02JyEAECBAgQIECAAAECBAgQIECAAAECBEZKoMsfxvzlyNQb3h/fgNCxzb6r+o65ptk+Kv0LzbjWE69Wa4w/Pz4aGzs7/WHJacljzZyOAAECBAgQIECAAAECBAgQIECAAAECBAhMKNDVD2PWeerHMO9NaomTn0i2SR5IquD99aTazsmLyX210dNqGZVq2yfXjY/WL5fymYyPTm5KDkrOS65Irk40AgQIECBAgAABAgQIECBAgAABAgQIECAwKwJVvK4lTNp8L+N6K7y2a/3uI5Nq9fb2k+Ojl//nuGzWsRc305ekr2J5zS1Pam3xKrBXsXzL5ODkhGTPZFrNcijTYnIQAQIECBAgQIAAAQIECBAgQIAAAQIECEwgsGvmqmD9cPIzSa3dXUut/Mek5v8mqfZosmp89PL/vCmbddxHeqY3z3jHZvvS9LVMyuHJJ5M6dnVSc29JpmyK4FMSOYAAAQIECBAgQIAAAQIECBAgQIAAAQIEJhGopU+qMH1H3/4qhNeb399p5r+Svo7btNluu5Oa+fe0Ez19LYdSn7kw2acZn5O+znFnsiKZsimCT0nkAAIECBAgQIAAAQIECBAgQIAAAQIECIycQFc/jPlUZOrN7N2S+iHLtq3NoFJLm1Rrf8zygPWbL/233X7wpZn1g+3S1RrgVew+P6llUapdn6xJbk72SHq/M5saAQIECBAgQIAAAQIECBAgQIAAAQIECBBYv2RJFw5V6L4rqeVLjuw5YY23Tf6xmbu26evN77YtzeDE5Jnkunay6S9Pv3VSx1fRu5Y/qbZkfTde/K75elNcI0CAAAECBAgQIECAAAECBAgQIECAAAECAxP4lZy5itG17ncta3JOUsug1Nwbkmq7JLU8Ss1dlpyZ3N5sX5S+t52cjTrujJ7Jeuu73iq/MjkweSC5MZmyWQ5lSiIHECBAgAABAgQIECBAgAABAgQIECBAgMAUAu/P/vrhyypeV1Ymxya9bVk2am3wenu8jqmi+QVJb6tieS2x0v9meB1zbtK+/V3n378mp2qK4FMJ2U+AAAECBAgQIECAAAECBAgQIECAAAEC0xGoH6ysN7brxzI31LbKzl0nOaD2HZ9Mdo7Nsm/nST474bQi+IQsJgkQIECAAAECBAgQIECAAAECBAgQIEBgFAQUwUfhKboHAgQIECBAgAABAgQIECBAgAABAgQIzExgo5kd7mgCBAgQIECAAAECBAgQIECAAAECBAgQIDB/BBTB58+zcqUECBAgQIAAAQIECBAgQIAAAQIECBAgMEMBRfAZgjmcAAECBAgQIECAAAECBAgQIECAAAECBOaPgCL4/HlWrpQAAQIECBAgQIAAAQIECBAgQIAAAQIEZiigCD5DMIcTIECAAAECBAgQIECAAAECBAgQIECAwPwRUASfP8/KlRIgQIAAAQIECBAgQIAAAQIECBAgQIDADAUUwWcI5nACBAgQIECAAAECBAgQIECAAAECBAgQmD8CiuDz51m5UgIECBAgQIAAAQIECBAgQIAAAQIECBCYoYAi+AzBHE6AAAECBAgQIECAAAECBAgQIECAAAEC80dAEXz+PCtXSoAAAQIECBAgQIAAAQIECBAgQIAAAQIzFFAEnyGYwwkQIECAAAECBAgQIECAAAECBAgQIEBg/ggogs+fZ+VKCRAgQIAAAQIECBAgQIAAAQIECBAgQGCGAorgMwRzOAECBAgQIECAAAECBAgQIECAAAECBAjMHwFF8PnzrFwpAQIECBAgQIAAAQIECBAgQIAAAQIECMxQQBF8hmAOJ0CAAAECBAgQIECAAAECBAgQIECAAIH5I6AIPn+elSslQIAAAQIECBAgQIAAAQIECBAgQIAAgRkKKILPEMzhBAgQIECAAAECBAgQIECAAAECBAgQIDB/BBTB58+zcqUECBAgQIAAAQIECBAgQIAAAQIECBAgMEMBRfAZgjmcAAECBAgQIECAAAECBAgQIECAAAECBOaPgCL4/HlWrpQAAQIECBAgQIAAAQIECBAgQIAAAQIEZiigCD5DMIcTIECAAAECBAgQIECAAAECBAgQIECAwPwRUASfP8/KlRIgQIAAAQIECBAgQIAAAQIECBAgQIDADAUUwWcI5nACBAgQIECAAAECBAgQIECAAAECBAgQmD8CiuDz51m5UgIECBAgQIAAAQIECBAgQIAAAQIECBCYoYAi+AzBHE6AAAECBAgQIECAAAECBAgQIECAAAEC80dAEXz+PCtXSoAAAQIECBAgQIAAAQIECBAgQIAAAQIzFFAEnyGYwwkQIECAAAECBAgQIECAAAECBAgQIEBg/ggogs+fZ+VKCRAgQIAAAQIECBAgQIAAAQIECBAgQGCGAorgMwRzOAECBAgQIECAAAECBAgQIECAAAECBAjMH4FNOrzUrXOufzPB+Z7K3Ff75hdne1myd3JPsjxZk7StivOHJjsldyQPJf3trEysSi7t32GbAAECBAgQIECAAAECBAgQIECAAAECBAh0LfCRnHDdBLm/74sOyfa3+45bme3de477XMZrk9XJc8kxSW87NRv1XW/rndzQeMWKFXW8RoAAAQIECBAgQIAAAQIECBAgQIAAAQIE/kUCl+RTVWh+R3JKT34247btkMF3k2eTM5IDk48l9bm7kmpHJLV9erI0qQL5LUnb9srg6eTadmI6vSL4dJQcQ4AAAQIECBAgQIAAAQIECBAgQIAAAQKTCbRF8C0nOyDzZyZV4H5f3zG3NfP7pq/idx3z6qTap5JHxkdjYxunvzV5PNmxmZtWpwg+LSYHESBAgAABAgQIECBAgAABAgQIECBAYKQEZvuHMY9t9K7qU7ym2T4q/QvNeEnTL0r/fDM+O/1hyWnJY82cjgABAgQIECBAgAABAgQIECBAgAABAgQITCjQ5Q9jtl/wgQxquZNvJF9Men8Uc+dsv5jcl/S2dt3w7TN5XbOjlkv5THJ0clNyUHJeckVydaIRIECAAAECBAgQIECAAAECBAgQIECAAIFZE7go31TLmNSb3NVXquBdy6S0rd7efrLd6OmPy7iOv7iZq8/UZ2tuebJfcm9SxfJabuXg5IRkz2RazXIo02JyEAECBAgQIECAAAECBAgQIECAAAECBAhMIlBvlW/R7HtF+uOTbydVyH5TUu3RZNX46OX/qf113Ed6pjfPuF33+9KMq7h+ePLJpI5dndTcW5IpmyL4lEQOIECAAAECBAgQIECAAAECBAgQIECAAIEZCrwrx1fB+j83n/tKs71ps912JzXz72knevpaDqXOcWGyTzM+J32d485kRTJlUwSfksgBBAgQIECAAAECBAgQIECAAAECBAgQGDmBQf8wZr0JXq1+3LJa+2OWB6zffOm/7faDL82sH2yXrtYAr2L3+cl+SbXrkzXJzckeSXv+DDUCBAgQIECAAAECBAgQIECAAAECBAgQILBeoMsi+BYToP77Zu4fmv7apq83v9u2NIMTk2eS9kcx232XZ7B1UsdX0buWP6m2ZH03Xvyu+XpTXCNAgAABAgQIECBAgAABAgQIECBAgAABAgMRqGVKaq3vP0nekdQyKH+bVHH6y8nipNouSf0wZs1flpyZ3N5s1w9r9raTs1HHndEzWW991w9mXpkcmDyQ3JhM2SyHMiWRAwgQIECAAAECBAgQIECAAAECBAgQIEBgEoFXZv7zSftWdhWvv5fUj1julPS2ZdmotcHXJnXco8kFSW+rYvlTSf+b4XXMuUn7PSsz3r8mp2qK4FMJ2U+AAAECBAgQIECAAAECBAgQIECAAAECUwlsngN+JNktmWqpla1yzK7JRK32HZ9sM9HOzG2W7DzJvgmnFcEnZDFJgAABAgQIECBAgAABAgQIECBAgAABAqMgoAg+Ck/RPRAgQIAAAQIECBAgQIAAAQIECBAgQGBmAlO9rT2zszmaAAECBAgQIECAAAECBAgQIECAAAECBAgMkYAi+BA9DJdCgAABAgQIECBAgAABAgQIECBAgAABAt0KKIJ36+lsBAgQIECAAAECBAgQIECAAAECBAgQIDBEAorgQ/QwXAoBAgQIECBAgAABAgQIECBAgAABAgQIdCugCN6tp7MRIECAAAECBAgQIECAAAECBAgQIECAwBAJKIIP0cNwKQQIECBAgAABAgQIECBAgAABAgQIECDQrYAieLeezkaAAAECBAgQIECAAAECBAgQIECAAAECQySgCD5ED8OlECBAgAABAgQIECBAgAABAgQIECBAgEC3Aorg3Xo6GwECBAgQIECAAAECBAgQIECAAAECBAgMkYAi+BA9DJdCgAABAgQIECBAgAABAgQIECBAgAABAt0KKIJ36+lsBAgQIECAAAECBAgQIECAAAECBAgQIDBEAorgQ/QwXAoBAgQIECBAgAABAgQIECBAgAABAgQIdCugCN6tp7MRIECAAAECBAgQIECAAAECBAgQIECAwBAJKIIP0cNwKQQIECBAgAABAgQIECBAgAABAgQIECDQrYAieLeezkaAAAECBAgQIECAAAECBAgQIECAAAECQySgCD5ED8OlECBAgAABAgQIECBAgAABAgQIECBAgEC3Aorg3Xo6GwECBAgQIECAAAECBAgQIECAAAECBAgMkYAi+BA9DJdCgAABAgQIECBAgAABAgQIECBAgAABAt0KKIJ36+lsBAgQIECAAAECBAgQIECAAAECBAgQIDBEAorgQ/QwXAoBAgQIECBAgAABAgQIECBAgAABAgQIdCugCN6tp7MRIECAAAECBAgQIECAAAECBAgQIECAwBAJKIIP0cNwKQQIECBAgAABAgQIECBAgAABAgQIECDQrYAieLeezkaAAAECBAgQIECAAAECBAgQIECAAAECQySgCD5ED8OlECBAgAABAgQIECBAgAABAgQIECBAgEC3Aorg3Xo6GwECBAgQIECAAAECBAgQIECAAAECBAgMkYAi+BA9DJdCgAABAgQIECBAgAABAgQIECBAgAABAt0KKIJ36+lsBAgQIECAAAECBAgQIECAAAECBAgQIDBEAorgQ/QwXAoBAgQIECBAgAABAgQIECBAgAABAgQIdCugCN6tp7MRIECAAAECBAgQIECAAAECBAgQIECAwBAJbDKga1mU8/7bZPPky8kTSW9bnI1lyd7JPcnyZE3StirOH5rslNyRPJT0t7MysSq5tH+HbQIECBAgQIAAAQIECBAgQIAAAQIECBAgMEiB03PydU3e3/dFh2T72z3767iVye5J2z6XwdpkdfJcckzS207NRn3ubb2TGxqvWLGijtcIECBAgAABAgQIECBAgAABAgQIECBAgMAPJbBtPl1vflfRudJbBN8h299Nnk3OSA5MPpbUcXcl1Y5IarsK6UuTKpDfkrRtrwyeTq5tJ6bTK4JPR8kxBAgQIECAAAECBAgQIECAAAECBAgQIDCVwB/ngCpiX9j0vUXwM5u596Xvbbdloz6zb9K+Rf7qjKt9KnlkfDQ2tnH6W5PHkx2buWl1iuDTYnIQAQIECBAgQIAAAQIECBAgQIAAAQIERkqg6x/GrHW+q4j96eSLE0gd28xd1bfvmmb7qPQvNOMlTV/riz/fjM9Of1hyWvJYM6cjQIAAAQIECBAgQIAAAQIECBAgQIAAAQITCnRZBK9idf1IZf1Y5X+a8NvGxnbO/IvJfX3772+2t09/dzOu5VKq4H10Um+KH5Scl1yRXJ1oBAgQIECAAAECBAgQIECAAAECBAgQIEBg1gR+Nd9US5q8p/nGKl7Xdu9yKPX29pPN/t7uuGzUsRc3k5ekr2J5zS1P9kvuTapYvmVycHJCsmcyrWY5lGkxOYgAAQIECBAgQIAAAQIECBAgQIAAAQIEJhB4ZeZq3e6vJps2+ycqgj+affWmeH97Uyaq4P2Rnh2bZ9yu+11vmNcyKYcnn0zq2NVJzb0lmbIpgk9J5AACBAgQIECAAAECBAgQIECAAAECBAiMnEBXy6G8OzI7JdcnP51UUfvQpNo+yc8mmyRPJEuTtlCe4Xjbpum/1fTVfS+pN8ermF5Lo3w4+U5ySlJvl2+V/EPygUQjQIAAAQIECBAgQIAAAQIECBAgQIAAAQIDE/jvOXO9nb2h/Gj239gc8+Ppe1u9AV6frSVOett22Xg4+VJShfPaX8fV+uDVfj+pH82s9cg32LwJvkEeOwkQIECAAAECBAgQIECAAAECBAgQIDCSAl29CX52dH6iL+9rxP4k/ZHJyuTaZu6kpq+u3gw/MXkmuS7pbZdnY+ukjl+TvJBUW7K+Gy9+13wVxjUCBAgQIECAAAECBAgQIECAAAECBAgQIDBrArWMSRWn39/zjbtkXD+MWfOXJWcmtzfbF6XvbSdno46rpVDatkcGLyZXJgcmDyT1dvmUzZvgUxI5gAABAgQIECBAgAABAgQIECBAgAABAgRmIPDvcmwVsds3wtuPLsvgK8napPY/mlyQ9LYqlj+V9L8ZXsecm7Rvf6/MeP+anKopgk8lZD8BAgQIECBAgAABAgQIECBAgAABAgQIzFTgFfnAZOt11w9b7jrJCWvf8Un7g5n9h22WiZ37Jze0rQi+IR37CBAgQIAAAQIECBAgQIAAAQIECBAgMJoCmwz4tmqd78laveldmajV/FUT7Wjmnkv/yAb220WAAAECBAgQIECAwAwFfuq9X3j9DD/icAIEhkTghouPvHlILsVlECBAgACBoROY7C3tobvQH/aC6k3wvffee8Hc7w/r5fMECBAgQIAAAQILS2C8AL520U0L667dLYEREtho3RsUwkfoeboVAgQIEOhUYKNOz+ZkBAgQIECAAAECBAgQIECAAAECBAgQIEBgiAQUwYfoYbgUAgQIECBAgAABAgQIECBAgAABAgQIEOhWQBG8W09nI0CAAAECBAgQIECAAAECBAgQIECAAIEhElAEH6KH4VIIECBAgAABAgQIECBAgAABAgQIECBAoFsBRfBuPZ2NAAECBAgQIECAAAECBAgQIECAAAECBIZIQBF8iB6GSyFAgAABAgQIECBAgAABAgQIECBAgACBbgUUwbv1dDYCBAgQIECAAAECBAgQIECAAAECBAgQGCIBRfAhehguhQABAgQIECBAgAABAgQIECBAgAABAgS6FVAE79bT2QgQIECAAAECBAgQIECAAAECBAgQIEBgiAQUwYfoYbgUAgQIECBAgAABAgQIECBAgAABAgQIEOhWQBG8W09nI0CAAAECBAgQIECAAAECBAgQIECAAIEhElAEH6KH4VIIECBAgAABAgQIECBAgAABAgQIECBAoFsBRfBuPZ2NAAECBAgQIECAAAECBAgQIECAAAECBIZIQBF8iB6GSyFAgAABAgQIECBAgAABAgQIECBAgACBbgUUwbv1dDYCBAgQIECAAAECBAgQIECAAAECBAgQGCIBRfAhehguhQABAgQIECBAgAABAgQIECBAgAABAgS6FVAE79bT2QgQIECAAAECBAgQIECAAAECBAgQIEBgiAQUwYfoYbgUAgQIECBAgAABAgQIECBAgAABAgQIEOhWQBG8W09nI0CAAAECBAgQIECAAAECBAgQIECAAIEhElAEH6KH4VIIECBAgAABAgQIECBAgAABAgQIECBAoFsBRfBuPZ2NAAECBAgQIECAAAECBAgQIECAAAECBIZIQBF8iB6GSyFAgAABAgQIECBAgAABAgQIECBAgACBbgUUwbv1dDYCBAgQIECAAAECBAgQIECAAAECBAgQGCIBRfAhehguhQABAgQIECBAgAABAgQIECBAgAABAgS6FVAE79bT2QgQIECAAAECBAgQIECAAAECBAgQIEBgiAQUwYfoYbgUAgQIECBAgAABAgQIECBAgAABAgQIEOhWQBG8W09nI0CAAAECBAgQIECAAAECBAgQIECAAIEhElAEH6KH4VIIECBAgAABAgQIECBAgAABAgQIECBAoFsBRfBuPZ2NAAECBAgQIECAAAECBAgQIECAAAECBIZIYJOOr2WrnO+IZMfk4eTvkqeS/rY4E8uSvZN7kuXJmqRtVZw/NNkpuSN5KOlvZ2ViVXJp/w7bBAgQIECAAAECBAgQIECAAAECBAgQIECgBLosgp+Z8/1esrRO3LQn0h+ffKGdSH9Icl2ybc/c/Rm/Mfl6M/cX6d+cPJtsnPxC8ldJ207N4MPJb7YTegIECBAgQIAAAQIECBAgQIAAAQIECBAg0C/Q5XIob83J/zo5NqlC96eT7ZNPJG3bIYMbki2SdyQHJZckeyVV+K5Wb5Ifl1SBuz7/UHJO0rY6tj5TRfHL2kk9AQIECBAgQIAAAQIECBAgQIAAAQIECBDoF+jyTfCfzskf7/mCX8/4l5L9ko2TF5MTky2Ts5O2OH5nxocnr0v2Ter4atcnq5Nbkjp3tTrPlUm9IX5KohEgQIAAAQIECBAgQIAAAQIECBAgQIAAgUkFunwTvLcAXl9Yxe5NkyeTKoBXq7fEq121vnvpv9c0o6PSv9CMlzT9ovTPN+Mqnh+WnJY81szpCBAgQIAAAQIECBAgQIAAAQIECBAgQIDAhAJdFsHbL6jC948lf57U+f9L0radM6iC+H3tRNPXmuDVavmTu8dHY2NnpK+C99HJbUktnXJeckVydaIRIECAAAECBAgQIECAAAECBAgQIECAAIENCnS5HEp90f9Ndun5xv+T8fk926/K+Ome7Xa4qhlsnb6WR/l4Uj+0WeuG/2PyoeSzyUPJO5ODk3+dfDl5INEIECBAgAABAgQIECBAgAABAgQIECBAgMAPCHRdBK9idRW6/1Xy88m/S25K3pjUkibrksVJf2uXS2n7d+WA30lemdSyJ5cmr0mOSqpAfkryvaTO9WvJZxKNAAECBAi8JPBT7/3C61/aMCBAgACBqQXWjr1+6oMcQYAAAQIECBAgQGD+CXRdBL+8h+C9Gf9dckTy1qT2PZHsmNSSKWuStm3TDL7VTqSvInellkOppVF+N/lOUgXw9ycfTWqZlA8kiuBB0AgQIEBgvcB4AXztovr/hNUIECBAgAABAgQIECBAgACBBS6w0QDv/5mc+8+a89db3NXaH7M8YP3mS/9ttx98aWb9YLt0VyS1REotq7JfUu36pIroNyd7JIsSjQABAgQIECBAgAABAgQIECBAgAABAgQIvEygqyJ4vdldP4bZ3w5pJlY2/bVNf1LPgUszPjGpovl1PfM1rLfHt07q+Cp6v5BUW7K+Gy9+13wts6IRIECAAAECBAgQIECAAAECBAgQIECAAIGXCXS1HMruOevdyd8mVciuH7qsZUyOSR5J6kctq/1pcm5Sa35X8furSRW4d0t+L1mdtO3kDE5I6scxv9ZMLk+/Njk9qaVSat3xWhJFI0CAAAECBAgQGCmBdeePbTT+v/obqbtyMwQIEBiUwA0XH3nzoM7tvAQIECBAYL4LdFUEfzAQ/yV5S1I/XlmtitX/O6mCd60FXu3h5CeTK5PfSGoZk1oi5cKkiuNt2yWDP0w+n3yinUz/z8kHk/OSKp7fn/x2ohEgQIAAAQIECIySQArgCjqj9EDdCwECBAgQIECAAIG5E+iqCF7LlLw9qbe2d03qvN9Mnk/625czsX+yVbJlUsf1t3qT/K3JTf07sn1B8tFk26TeMtcIECBAgAABAgQIECBAgAABAgQIECBAgMCEAl0VwduT19vf32g3puifyv7KRK3mr5poRzP3XHoF8A0A2UWAAAECBAgQIECAAAECBAgQIECAAAECY1lpUSNAgAABAgQIECBAgAABAgQIECBAgAABAiMqoAg+og/WbREgQIAAAQIECBAgQIAAAQIECBAgQICAN8H9DRAgQIAAAQIECBAgQIAAAQIECBAgQIDACAt4E3yEH65bI0CAAAECBAgQIECAAAECBAgQIECAwEIXUARf6H8B7p8AAQIECBAgQIAAAQIECBAgQIAAAQIjLKAIPsIP160RIECAAAECBAgQIECAAAECBAgQIEBgoQsogi/0vwD3T4AAAQIECBAgQIAAAQIECBAgQIAAgREWUAQf4Yfr1ggQIECAAAECBAgQIECAAAECBAgQILDQBRTBF/pfgPsnQIAAAQIECBAgQIAAAQIECBAgQIDACAsogo/ww3VrBAgQIECAAAECBAgQIECAAAECBAgQWOgCiuAL/S/A/RMgQIAAAQIECBAgQIAAAQIECBAgQGCEBRTBR/jhujUCBAgQIECAAAECBAgQIECAAAECBAgsdAFF8IX+F+D+CRAgQIAAAQIECBAgQIAAAQIECBAgMMICiuAj/HDdGgECBAgQIECAAAECBAgQIECAAAECBBa6gCL4Qv8LcP8ECBAgQIAAAQIECBAgQIAAAQIECBAYYQFF8BF+uG6NAAECBAgQIECAAAECBAgQIECAAAECC11AEXyh/wW4fwIECBAgQIAAAQIECBAgQIAAAQIECIywgCL4CD9ct0aAAAECBAgQIECAAAECBAgQIECAAIGFLqAIvtD/Atw/AQIECBAgQIAAAQIECBAgQIAAAQIERlhAEXyEH65bI0CAAAECBAgQIECAAAECBAgQIECAwEIXUARf6H8B7p8AAQIECBAgQIAAAQIECBAgQIAAAQIjLKAIPsIP160RIECAAAECBAgQIECAAAECBAgQIEBgoQsogi/0vwD3T4AAAQIECBAgQIAAAQIECBAgQIAAgREWUAQf4Yfr1ggQIECAAAECBAgQIECAAAECBAgQILDQBRTBF/pfgPsnQIAAAQIECBAgQIAAAQIECBAgQIDACAsogo/ww3VrBAgQIECAAAECBAgQIECAAAECBAgQWOgCiuAL/S/A/RMgQIAAAQIECBAgQIAAAQIECBAgQGCEBRTBR/jhujUCBAgQIECAAAECBAgQIECAAAECBAgsdAFF8IX+F+D+CRAgQIAAAQIECBAgQIAAAQIECBAgMMICm3R8b6/I+Y5MdkoeSv4+eTrpb4szsSzZO7knWZ6sSdpWxflDkzrPHUmdq7+dlYlVyaX9O2wTIECAAAECBAgQIECAAAECBAgQIECAAIES6LIIfkbO9+GkCuFtezSDX0xuaSfSH5Jcl2zbM3d/xm9Mvt7M/UX6NyfPJhsnv5D8VdK2UzOo7/rNdkJPgAABAgQIECBAgAABAgQIECBAgAABAgT6BbpcDuX0nPwrySnJzyV/ndSb3P8tadsOGdyQbJG8IzkouSTZK6nCd7UjkuOSKnBvn9Rb4Ockbatj6zNVFL+sndQTIECAAAECBAgQIECAAAECBAgQIECAAIF+gS7fBK+3tR9I1jVf8jfpv5X8SLJj8lhyYrJlcnbyiaTancnhyeuSfZP9kmrXJ6uTeov8p5Nq9Vb4lUm9IV7Fdo0AAQIECBAgQIAAAQIECBAgQIAAAQIECEwq0OWb4LWkSVsAry98KqkidrW167uxY5v+qqZvu2uawVHpX2jGS5p+Ufrnm3EVzw9LTkuqqK4RIECAAAECBAgQIECAAAECBAgQIECAAIFJBbosgvd/Sf3w5dbJN5Inmp07p38xua/ZbrsqoFer5U/uHh+NjdUa41XwPjq5LamlU85LrkiuTjQCBAgQIECAAAECBAgQIECAAAECBAgQILBBgUEWwS9svvmP07dviL8q46cnuKJVzVwVzWt5lI8nb09uTR5NPpTUMigPJe9MDk5OSPZMNAIECBAgQIAAAQIECBAgQIAAAQIECBAgMKHAoIrg9Rb3Mck/JB/r+eYqhi/u2W6H9XZ4tbZ/V8avSOqHNV+b1I9kvib51aQK5LcnVRSvN8rfkmgECBAgQIAAAQIECBAgQIAAAQIECBAgQOAHBAZRBP/ZfMsfJI8kP5esSdpWy6IsTTZtJ5p+m6b/Vs/89zKudb9rOZQqqn84+U5ySvL+ZKukiuwfSDQCBAgQIECAAAECBAgQIECAAAECBAgQIPADAl0XwY/IN3w2qeVN3pR8Pelt7Y9ZHtA7mXG7/WDf/HbZrjXAa4mU85P9kmrXJ1VcvznZI6kfz9QIECBAgAABAgQIECBAgAABAgQIECBAgMDLBLosgh+YM1/bnL0K4F9+2Tet32j3n9Szr94MPzF5JrmuZ76Glye1TngdX0XvF5JqS9Z348Xvmm/XHG+mdQQIECBAgAABAgQIECBAgAABAgQIECBAYGxsk44QNs55rklemdyT/HKTdOPtC/nvnyV/mpyb1JrfVfz+alIF7t2S30tWJ207OYP68ct3JF9rJpenX5ucntRyKT+f3JZoBAgQIECAAAECBAgQIECAAAECBAgQIEDgBwS6KoLXeWqN7mq1ZEm7bMn4RP5TP3JZRfCHk59M6kctfyOpZQKxFg4AAD9KSURBVExqiZQLkyqOt22XDP4w+XzyiXYy/T8nH0zOS6p4fn/y24lGgAABAgQIECBAgAABAgQIECBAgAABAgR+QKCrIvhzOXO92T2dVsuk7J9U0XzL5JtJf1uVibcmN/XvyPYFyUeTbZP68U2NAAECBAgQIECAAAECBAgQIECAAAECBAhMKNBVEXzCk08x+VT2VyZqNX/VRDuauSq6K4BvAMguAgQIECBAgAABAgQIECBAgAABAgQIEBgb6/KHMXkSIECAAAECBAgQIECAAAECBAgQIECAAIGhElAEH6rH4WIIECBAgAABAgQIECBAgAABAgQIECBAoEsBRfAuNZ2LAAECBAgQIECAAAECBAgQIECAAAECBIZKQBF8qB6HiyFAgAABAgQIECBAgAABAgQIECBAgACBLgUUwbvUdC4CBAgQIECAAAECBAgQIECAAAECBAgQGCoBRfChehwuhgABAgQIECBAgAABAgQIECBAgAABAgS6FFAE71LTuQgQIECAAAECBAgQIECAAAECBAgQIEBgqAQUwYfqcbgYAgQIECBAgAABAgQIECBAgAABAgQIEOhSQBG8S03nIkCAAAECBAgQIECAAAECBAgQIECAAIGhElAEH6rH4WIIECBAgAABAgQIECBAgAABAgQIECBAoEsBRfAuNZ2LAAECBAgQIECAAAECBAgQIECAAAECBIZKQBF8qB6HiyFAgAABAgQIECBAgAABAgQIECBAgACBLgUUwbvUdC4CBAgQIECAAAECBAgQIECAAAECBAgQGCoBRfChehwuhgABAgQIECBAgAABAgQIECBAgAABAgS6FFAE71LTuQgQIECAAAECBAgQIECAAAECBAgQIEBgqAQUwYfqcbgYAgQIECBAgAABAgQIECBAgAABAgQIEOhSQBG8S03nIkCAAAECBAgQIECAAAECBAgQIECAAIGhElAEH6rH4WIIECBAgAABAgQIECBAgAABAgQIECBAoEsBRfAuNZ2LAAECBAgQIECAAAECBAgQIECAAAECBIZKQBF8qB6HiyFAgAABAgQIECBAgAABAgQIECBAgACBLgUUwbvUdC4CBAgQIECAAAECBAgQIECAAAECBAgQGCoBRfChehwuhgABAgQIECBAgAABAgQIECBAgAABAgS6FFAE71LTuQgQIECAAAECBAgQIECAAAECBAgQIEBgqAQUwYfqcbgYAgQIECBAgAABAgQIECBAgAABAgQIEOhSQBG8S03nIkCAAAECBAgQIECAAAECBAgQIECAAIGhElAEH6rH4WIIECBAgAABAgQIECBAgAABAgQIECBAoEsBRfAuNZ2LAAECBAgQIECAAAECBAgQIECAAAECBIZKQBF8qB6HiyFAgAABAgQIECBAgAABAgQIECBAgACBLgUUwbvUdC4CBAgQIECAAAECBAgQIECAAAECBAgQGCqBQRXB98ldvjlZOsndLs7865JfSw5MNk16W13X4cnxyW69O3rGZ2X8jp5tQwIECBAgQIAAAQIECBAgQIAAAQIECBAg8DKBrovgm+fsH07+Mflfyb5JfzskE48ktyWfSr6U3JvsnrTtLzL4YvLpZGVyTNLbTs1Gfc8LvZPGBAgQIECAAAECBAgQIECAAAECBAgQIECgV6DLIviP5cTLk3pDe03vl/SMd8j4hmSLpN7iPii5JNkrqcJ3tSOS45LfTLZPHkrOSdpWx9Zn/iq5rJ3UEyBAgAABAgQIECBAgAABAgQIECBAgACBfoEui+DH5uTPJT+ZfKb/i5rtE9NvmXww+URyZ/Lu5O+TZUm9Ob5fUu36ZHVyS7JnUm3j5Mrk2eSURCNAgAABAgQIECBAgAABAgQIECBAgAABApMKdFkE//18y2uTmyb9trGxKpRXu2p999J/r2lGR6VvlzhZ0swtSv98Mz47/WHJacljzZyOAAECBAgQIECAAAECBAgQIECAAAECBAhMKNBlEbzezl474bd8f3LnDF9M7vv+1Pjo/ma7lj+5uxmfkb4K3kcntX54LZ1yXnJFcnWiESBAgAABAgQIECBAgAABAgQIECBAgACBDQp0WQTf4Bc1O1+V/ukJDlzVzG2dvpZI+Xjy9uTW5NHkQ0ktg/JQ8s7k4OSEZM9EI0CAAAECBAgQIECAAAECBAgQIECAAAECEwpsMuHs4CbX5dSLJzh9vR1ere3flfHvJK9MatmTS5PXJLVcShXIT0m+l9S5fi2ZbA3y7Jr/7afe+4XXz/+7cAcECBCYRYG1Y6+fxW/zVQQIECBAgAABAgQIECBAgMAQC8x2EfyJWOyYbJqs6XHZphl/q2euityVWg6llkb53eQ7SRXA3598NKllUj6QjGwRfLwAvnbRhtZZz+1rBAgQIECAAAECBAgQIECAAAECBAgQIDCRwEYTTQ5wrv0xywP6vqPdfrBvfrts1xrgtUTK+cl+SbXrkyqi35zskdSPZ2oECBAgQIAAAQIECBAgQIAAAQIECBAgQOBlArNdBL+2+faTeq5iacYnJs8k1/XM1/DypNYJr+Or6P1CUm3J+m68+F3ztcyKRoAAAQIECBAgQIAAAQIECBAgQIAAAQIEXibQ5XIor82Zj2/OfmDTvy39N5O7k6uTP03OTWrN7yp+fzWpAvduye8lq5O2nZxB/fjlO5KvNZPL069NTk9qqZSfT2pJFI0AAQJDLLDu/LGNxv+XK0N8jS6NAAECwyVww8VH3jxcV+RqCBAgQIAAAQIECBCYrwJdFsEPC0Ktz93bfr3ZqOJ3FcEfTn4yuTL5jaSWMaklUi5Mqjjetl0y+MPk88kn2sn0/5x8MDkvqeL5/clvJxoBAgSGVyAFcMWc4X08rowAAQIECBAgQIAAAQIECBAYbYEui+C1dEllqvblHLB/slWyZfLNpL+tysRbk5v6d2T7gqR+FHPb5JFEI0CAAAECBAgQIECAAAECBAgQIECAAAECEwp0WQSf8As2MPlU9lUmajV/1UQ7mrnn0iuAbwDILgIECBAgQIAAAQIECBAgQIAAAQIECBAYyyq1GgECBAgQIECAAAECBAgQIECAAAECBAgQGFEBRfARfbBuiwABAgQIECBAgAABAgQIECBAgAABAgS8Ce5vgAABAgQIECBAgAABAgQIECBAgAABAgRGWMCb4CP8cN0aAQIECBAgQIAAAQIECBAgQIAAAQIEFrqAIvhC/wtw/wQIECBAgAABAgQIECBAgAABAgQIEBhhAUXwEX64bo0AAQIECBAgQIAAAQIECBAgQIAAAQILXUARfKH/Bbh/AgQIECBAgAABAgQIECBAgAABAgQIjLCAIvgIP1y3RoAAAQIECBAgQIAAAQIECBAgQIAAgYUuoAi+0P8C3D8BAgQIECBAgAABAgQIECBAgAABAgRGWEARfIQfrlsjQIAAAQIECBAgQIAAAQIECBAgQIDAQhdQBF/ofwHunwABAgQIECBAgAABAgQIECBAgAABAiMsoAg+wg/XrREgQIAAAQIECBAgQIAAAQIECBAgQGChCyiCL/S/APdPgAABAgQIECBAgAABAgQIECBAgACBERZQBB/hh+vWCBAgQIAAAQIECBAgQIAAAQIECBAgsNAFFMEX+l+A+ydAgAABAgQIECBAgAABAgQIECBAgMAICyiCj/DDdWsECBAgQIAAAQIECBAgQIAAAQIECBBY6AKK4Av9L8D9EyBAgAABAgQIECBAgAABAgQIECBAYIQFFMFH+OG6NQIECBAgQIAAAQIECBAgQIAAAQIECCx0AUXwhf4X4P4JECBAgAABAgQIECBAgAABAgQIECAwwgKK4CP8cN0aAQIECBAgQIAAAQIECBAgQIAAAQIEFrqAIvhC/wtw/wQIECBAgAABAgQIECBAgAABAgQIEBhhAUXwEX64bo0AAQIECBAgQIAAAQIECBAgQIAAAQILXUARfKH/Bbh/AgQIECBAgAABAgQIECBAgAABAgQIjLCAIvgIP1y3RoAAAQIECBAgQIAAAQIECBAgQIAAgYUuoAi+0P8C3D8BAgQIECBAgAABAgQIECBAgAABAgRGWEARfIQfrlsjQIAAAQIECBAgQIAAAQIECBAgQIDAQhdQBF/ofwHunwABAgQIECBAgAABAgQIECBAgAABAiMsoAg+wg/XrREgQIAAAQIECBAgQIAAAQIECBAgQGChCyiCL/S/APdPgAABAgQIECBAgAABAgQIECBAgACBERZQBB/hh+vWCBAgQIAAAQIECBAg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b4IHQSNAgAABAgQIECBAYPACiuCDN/YNBAiMlsC2uZ37kirkfHK0bs3dECBAYKACVfR+a9IuJXXJQL/NyQkQIECAAAECBAgQINAIKIL7UyBAgMD0BerH3L6YVAH884llUIKgESBAYIYC787x9e9orQuuESBAgAABAgQIECBAYOACiuADJ/YFBAiMiMDi3Mf1SRVu/j55RaIRIECAwMwF3pyP1L+lj8/8oz5BgAABAgQIECBAgACBmQsogs/czCcIEFh4Ahvnlv9nUkWbu5NtEo0AAQIENizwxuw+uu+Q+vf0c0n9e3pj3z6bBAgQIECAAIGRFfA/Ix7ZR+vGhlhgz1zbRc31bdX0u6T/02Z8TfpPN2MdAQIECIyNnRSE4xuIKuB8tg9lTbZ/Pnm2b94mAQIEFrLAMbn5Wvqkfkfhb5Pnkjck+yYvJOckGgECBAgQIECAAAECBAYicETOWm/fTBY/UjQQdiclQGAeC9QPuU32b2bNVxF8i3l8fy6dAAECgxDYOyf9y6T+Pwjbf0PXZnxb8hOJRoAAAQIECBBYMAKLFsydulECBAgQIECAAAECBAgsPIHNcsu7JZsm/5z4X80EQSNAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgACB/9/evQBZUpUHAD5957E798KirAvuY2YEERDEBAwpQ9BENAI+CpCK0UqiBA1oiRqIiGhixahAIkY0xojxXYpGiUksJSUhPhJT0URMlKAVEXRnH4qLPJ07y87MPfm778ydubvz2mVnw26+rrrT3eecPn3O113l8t/jfwkQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAEC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WAlfrRI/PzyiaX5tlMV+CVtRfC9aPTtynv9l/FjkZ+OyO9PYprenVvGH4fuxVKvdEAbvjX6jXbEm6s+I/f3xhcSzltD7/tdkX9gurPKmqL42PtdHPvqPxL78MuO58ax/IR7tN+LZfDLOl77tzfks/a5aEiBAgAABAgQIECBAYEkCguBLYtKIAAECBA5YgaI4vZpbK1835xy3b/xBGQCO4OCpud54amreFYHgqdzgxc75qIupHN2TU/tIdhJbuWo6/hw33X9EjQ+N44+m5uTVud5zUgTKz47zt0Uu7tEI/l4Z4eXz4nwmIF/Upo+n+o3acivK+8T/lFe5yNtF8bfdZpexdepTarauz/XaKVFybszrgrImAp+Tkfrl02ly/NJZgf4dKU88JafeMqf5M6LtyWXDXBS3RWD/VWl00xfKa6e2he/bGr8m1/piBXk6NnKkXxq9/EV13faNH8+NoZ44vjLqLonV6FVxGERKmNZbpvqOOe6WQeeyWQdzjW+ushJjMp5Xu266g6IV5/E4c5n3fYGtmH4Hpvfztl267Xxzn+56PtuqfmocRU/3uJsj74v3Ol654i3xfl5YNo1n8rNw/1BqNl9dnrYvn8N97vEsfT5Vx/4QIECAAAECBAgQIECAAAECBAgQILCvBPpTWhcJkhfcItJcrgKP0Gh7i6DtmoPmuCLaHd6Yo3yRokc8Iq1Yd0w0iuurbWX87Zs6ntpV/U7Xz6rapXy+sc26pnPYl1YOHZlWrD86SsJhwa2/PcZDV83Tain3LdLA+g1p5drhuftYvzr1D0VO8sOmcpLv3GqXuU41mK+86/q5xjdXWXnRiqlPVwdTzzauWWzrelcWa1zWL2YbTRad43y28c4s9n6H94rBx8ZNpt/vckyztrnuPVdZ55IlzKfT1gEBAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQIECBAgAABAgQI7LbA/wKdSZJ9JyhjbQAAAABJRU5ErkJggg==', { base64: true })

			// C: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
			this.slideLayouts.forEach((layout, idx) => {
				zip.file(`ppt/slideLayouts/slideLayout${idx + 1}.xml`, genXml.makeXmlLayout(layout))
				zip.file(`ppt/slideLayouts/_rels/slideLayout${idx + 1}.xml.rels`, genXml.makeXmlSlideLayoutRel(idx + 1, this.slideLayouts))
			})
			this.slides.forEach((slide, idx) => {
				zip.file(`ppt/slides/slide${idx + 1}.xml`, genXml.makeXmlSlide(slide))
				zip.file(`ppt/slides/_rels/slide${idx + 1}.xml.rels`, genXml.makeXmlSlideRel(this.slides, this.slideLayouts, idx + 1))
				// Create all slide notes related items. Notes of empty strings are created for slides which do not have notes specified, to keep track of _rels.
				zip.file(`ppt/notesSlides/notesSlide${idx + 1}.xml`, genXml.makeXmlNotesSlide(slide))
				zip.file(`ppt/notesSlides/_rels/notesSlide${idx + 1}.xml.rels`, genXml.makeXmlNotesSlideRel(idx + 1))
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
			return await Promise.all(arrChartPromises).then(async () => {
				if (props.outputType === 'STREAM') {
					// A: stream file
					return await zip.generateAsync({ type: 'nodebuffer', compression: props.compression ? 'DEFLATE' : 'STORE' })
				} else if (props.outputType) {
					// B: Node [fs]: Output type user option or default
					return await zip.generateAsync({ type: props.outputType })
				} else {
					// C: Browser: Output blob as app/ms-pptx
					return await zip.generateAsync({ type: 'blob', compression: props.compression ? 'DEFLATE' : 'STORE' })
				}
			})
		})
	}

	// EXPORT METHODS

	/**
	 * Export the current Presentation to stream
	 * @param {WriteBaseProps} props - output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file stream
	 */
	async stream (props?: WriteBaseProps): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> {
		return await this.exportPresentation({
			compression: props?.compression,
			outputType: 'STREAM',
		})
	}

	/**
	 * Export the current Presentation as JSZip content with the selected type
	 * @param {WriteProps} props output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array>} file content in selected type
	 */
	async write (props?: WriteProps | WRITE_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Buffer | Uint8Array> {
		// DEPRECATED: @deprecated v3.5.0 - outputType - [[remove in v4.0.0]]
		const propsOutpType = typeof props === 'object' && props?.outputType ? props.outputType : props ? (props as WRITE_OUTPUT_TYPE) : null
		const propsCompress = typeof props === 'object' && props?.compression ? props.compression : false

		return await this.exportPresentation({
			compression: propsCompress,
			outputType: propsOutpType,
		})
	}

	/**
	 * Export the current Presentation. Writes file to local file system if `fs` exists, otherwise, initiates download in browsers
	 * @param {WriteFileProps} props - output file properties
	 * @returns {Promise<string>} the presentation name
	 */
	async writeFile (props?: WriteFileProps | string): Promise<string> {
		const fs = typeof require !== 'undefined' && typeof window === 'undefined' ? require('fs') : null // NodeJS
		// DEPRECATED: @deprecated v3.5.0 - fileName - [[remove in v4.0.0]]
		if (typeof props === 'string') console.log('Warning: `writeFile(filename)` is deprecated - please use `WriteFileProps` argument (v3.5.0)')
		const propsExpName = typeof props === 'object' && props?.fileName ? props.fileName : typeof props === 'string' ? props : ''
		const propsCompress = typeof props === 'object' && props?.compression ? props.compression : false
		const fileName = propsExpName ? (propsExpName.toString().toLowerCase().endsWith('.pptx') ? propsExpName : propsExpName + '.pptx') : 'Presentation.pptx'

		return await this.exportPresentation({
			compression: propsCompress,
			outputType: fs ? 'nodebuffer' : null,
		}).then(async content => {
			if (fs) {
				// Node: Output
				return await new Promise<string>((resolve, reject) => {
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
				return await this.writeFileToBrowser(fileName, content as Blob)
			}
		})
	}

	// PRESENTATION METHODS

	/**
	 * Add a new Section to Presentation
	 * @param {ISectionProps} section - section properties
	 * @example pptx.addSection({ title:'Charts' });
	 */
	addSection (section: SectionProps): void {
		if (!section) console.warn('addSection requires an argument')
		else if (!section.title) console.warn('addSection requires a title')

		const newSection: SectionProps = {
			_type: 'user',
			_slides: [],
			title: section.title,
		}

		if (section.order) this.sections.splice(section.order, 0, newSection)
		else this._sections.push(newSection)
	}

	/**
	 * Add a new Slide to Presentation
	 * @param {AddSlideProps} options - slide options
	 * @returns {PresSlide} the new Slide
	 */
	addSlide (options?: AddSlideProps): PresSlide {
		// TODO: DEPRECATED: arg0 string "masterSlideName" dep as of 3.2.0
		const masterSlideName = typeof options === 'string' ? options : options?.masterName ? options.masterName : ''
		let slideLayout: SlideLayout = {
			_name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
			_presLayout: this.presLayout,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slideNum: this.slides.length + 1,
		}

		if (masterSlideName) {
			const tmpLayout = this.slideLayouts.filter(layout => layout._name === masterSlideName)[0]
			if (tmpLayout) slideLayout = tmpLayout
		}

		const newSlide = new Slide({
			addSlide: this.addNewSlide,
			getSlide: this.getSlide,
			presLayout: this.presLayout,
			setSlideNum: this.setSlideNumber,
			slideId: this.slides.length + 256,
			slideRId: this.slides.length + 2,
			slideNumber: this.slides.length + 1,
			slideLayout,
		})

		// A: Add slide to pres
		this._slides.push(newSlide)

		// B: Sections
		// B-1: Add slide to section (if any provided)
		// B-2: Handle slides without a section when sections are already is use ("loose" slides arent allowed, they all need a section)
		if (options?.sectionTitle) {
			const sect = this.sections.filter(section => section.title === options.sectionTitle)[0]
			if (!sect) console.warn(`addSlide: unable to find section with title: "${options.sectionTitle}"`)
			else sect._slides.push(newSlide)
		} else if (this.sections && this.sections.length > 0 && (!options?.sectionTitle)) {
			const lastSect = this._sections[this.sections.length - 1]

			// CASE 1: The latest section is a default type - just add this one
			if (lastSect._type === 'default') lastSect._slides.push(newSlide)
			// CASE 2: There latest section is NOT a default type - create the defualt, add this slide
			else {
				this._sections.push({
					title: `Default-${this.sections.filter(sect => sect._type === 'default').length + 1}`,
					_type: 'default',
					_slides: [newSlide],
				})
			}
		}

		return newSlide
	}

	/**
	 * Create a custom Slide Layout in any size
	 * @param {PresLayout} layout - layout properties
	 * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
	 */
	defineLayout (layout: PresLayout): void {
		// @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
		if (!layout) console.warn('defineLayout requires `{name, width, height}`')
		else if (!layout.name) console.warn('defineLayout requires `name`')
		else if (!layout.width) console.warn('defineLayout requires `width`')
		else if (!layout.height) console.warn('defineLayout requires `height`')
		else if (typeof layout.height !== 'number') console.warn('defineLayout `height` should be a number (inches)')
		else if (typeof layout.width !== 'number') console.warn('defineLayout `width` should be a number (inches)')

		this.LAYOUTS[layout.name] = {
			name: layout.name,
			_sizeW: Math.round(Number(layout.width) * EMU),
			_sizeH: Math.round(Number(layout.height) * EMU),
			width: Math.round(Number(layout.width) * EMU),
			height: Math.round(Number(layout.height) * EMU),
		}
	}

	/**
	 * Create a new slide master [layout] for the Presentation
	 * @param {SlideMasterProps} props - layout properties
	 */
	defineSlideMaster (props: SlideMasterProps): void {
		if (!props.title) throw new Error('defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)')

		const newLayout: SlideLayout = {
			_margin: props.margin || DEF_SLIDE_MARGIN_IN,
			_name: props.title,
			_presLayout: this.presLayout,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slide: null,
			_slideNum: 1000 + this.slideLayouts.length + 1,
			_slideNumberProps: props.slideNumber || null,
			_slideObjects: [],
			background: props.background || null,
			bkgd: props.bkgd || null,
		}

		// STEP 1: Create the Slide Master/Layout
		genObj.createSlideMaster(props, newLayout)

		// STEP 2: Add it to layout defs
		this.slideLayouts.push(newLayout)

		// STEP 3: Add background (image data/path must be captured before `exportPresentation()` is called)
		if (props.background || props.bkgd) genObj.addBackgroundDefinition(props.background, newLayout)

		// STEP 4: Add slideNumber to master slide (if any)
		if (newLayout._slideNumberProps && !this.masterSlide._slideNumberProps) this.masterSlide._slideNumberProps = newLayout._slideNumberProps
	}

	// HTML-TO-SLIDES METHODS

	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} eleId - table HTML element ID
	 * @param {TableToSlidesProps} options - generation options
	 */
	tableToSlides (eleId: string, options: TableToSlidesProps = {}): void {
		// @note `verbose` option is undocumented; used for verbose output of layout process
		genTable.genTableToSlides(
			this,
			eleId,
			options,
			options?.masterSlideName ? this.slideLayouts.filter(layout => layout._name === options.masterSlideName)[0] : null
		)
	}
}


function getStyle1() {
	return `
	<?xml version="1.0"?>
<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="395">
  <cs:axisTitle>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:defRPr sz="1197"/>
  </cs:axisTitle>
  <cs:categoryAxis>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="15000"/>
            <a:lumOff val="85000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
    <cs:defRPr sz="1197"/>
  </cs:categoryAxis>
  <cs:chartArea mods="allowNoFillOverride allowNoLineOverride">
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:solidFill>
        <a:schemeClr val="bg1"/>
      </a:solidFill>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="15000"/>
            <a:lumOff val="85000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
    <cs:defRPr sz="1330"/>
  </cs:chartArea>
  <cs:dataLabel>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:defRPr sz="1197"/>
  </cs:dataLabel>
  <cs:dataLabelCallout>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="dk1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:spPr>
      <a:solidFill>
        <a:schemeClr val="lt1"/>
      </a:solidFill>
      <a:ln>
        <a:solidFill>
          <a:schemeClr val="dk1">
            <a:lumMod val="25000"/>
            <a:lumOff val="75000"/>
          </a:schemeClr>
        </a:solidFill>
      </a:ln>
    </cs:spPr>
    <cs:defRPr sz="1197"/>
    <cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1">
      <a:spAutoFit/>
    </cs:bodyPr>
  </cs:dataLabelCallout>
  <cs:dataPoint>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0">
      <cs:styleClr val="auto"/>
    </cs:fillRef>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:solidFill>
        <a:schemeClr val="phClr"/>
      </a:solidFill>
    </cs:spPr>
  </cs:dataPoint>
  <cs:dataPoint3D>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0">
      <cs:styleClr val="auto"/>
    </cs:fillRef>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:solidFill>
        <a:schemeClr val="phClr"/>
      </a:solidFill>
    </cs:spPr>
  </cs:dataPoint3D>
  <cs:dataPointLine>
    <cs:lnRef idx="0">
      <cs:styleClr val="auto"/>
    </cs:lnRef>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="28575" cap="rnd">
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
  </cs:dataPointLine>
  <cs:dataPointMarker>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0">
      <cs:styleClr val="auto"/>
    </cs:fillRef>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:solidFill>
        <a:schemeClr val="phClr"/>
      </a:solidFill>
      <a:ln w="9525">
        <a:solidFill>
          <a:schemeClr val="lt1"/>
        </a:solidFill>
      </a:ln>
    </cs:spPr>
  </cs:dataPointMarker>
  <cs:dataPointMarkerLayout symbol="circle" size="5"/>
  <cs:dataPointWireframe>
    <cs:lnRef idx="0">
      <cs:styleClr val="auto"/>
    </cs:lnRef>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="28575" cap="rnd">
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
  </cs:dataPointWireframe>
  <cs:dataTable>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="15000"/>
            <a:lumOff val="85000"/>
          </a:schemeClr>
        </a:solidFill>
      </a:ln>
    </cs:spPr>
    <cs:defRPr sz="1197"/>
  </cs:dataTable>
  <cs:downBar>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="dk1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:solidFill>
        <a:schemeClr val="dk1">
          <a:lumMod val="65000"/>
          <a:lumOff val="35000"/>
        </a:schemeClr>
      </a:solidFill>
      <a:ln w="9525">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="65000"/>
            <a:lumOff val="35000"/>
          </a:schemeClr>
        </a:solidFill>
      </a:ln>
    </cs:spPr>
  </cs:downBar>
  <cs:dropLine>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="35000"/>
            <a:lumOff val="65000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
  </cs:dropLine>
  <cs:errorBar>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="65000"/>
            <a:lumOff val="35000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
  </cs:errorBar>
  <cs:floor>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
  </cs:floor>
  <cs:gridlineMajor>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="15000"/>
            <a:lumOff val="85000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
  </cs:gridlineMajor>
  <cs:gridlineMinor>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="15000"/>
            <a:lumOff val="85000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
  </cs:gridlineMinor>
  <cs:hiLoLine>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="75000"/>
            <a:lumOff val="25000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
  </cs:hiLoLine>
  <cs:leaderLine>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="35000"/>
            <a:lumOff val="65000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
  </cs:leaderLine>
  <cs:legend>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:defRPr sz="1197"/>
  </cs:legend>
  <cs:plotArea mods="allowNoFillOverride allowNoLineOverride">
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
  </cs:plotArea>
  <cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride">
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
  </cs:plotArea3D>
  <cs:seriesAxis>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="15000"/>
            <a:lumOff val="85000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
    <cs:defRPr sz="1197"/>
  </cs:seriesAxis>
  <cs:seriesLine>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="9525" cap="flat">
        <a:solidFill>
          <a:srgbClr val="D9D9D9"/>
        </a:solidFill>
        <a:round/>
      </a:ln>
    </cs:spPr>
  </cs:seriesLine>
  <cs:title>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:defRPr sz="1862"/>
  </cs:title>
  <cs:trendline>
    <cs:lnRef idx="0">
      <cs:styleClr val="auto"/>
    </cs:lnRef>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:ln w="19050" cap="rnd">
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:prstDash val="sysDash"/>
      </a:ln>
    </cs:spPr>
  </cs:trendline>
  <cs:trendlineLabel>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:defRPr sz="1197"/>
  </cs:trendlineLabel>
  <cs:upBar>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="dk1"/>
    </cs:fontRef>
    <cs:spPr>
      <a:solidFill>
        <a:schemeClr val="lt1"/>
      </a:solidFill>
      <a:ln w="9525">
        <a:solidFill>
          <a:schemeClr val="tx1">
            <a:lumMod val="15000"/>
            <a:lumOff val="85000"/>
          </a:schemeClr>
        </a:solidFill>
      </a:ln>
    </cs:spPr>
  </cs:upBar>
  <cs:valueAxis>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1">
        <a:lumMod val="65000"/>
        <a:lumOff val="35000"/>
      </a:schemeClr>
    </cs:fontRef>
    <cs:defRPr sz="1197"/>
  </cs:valueAxis>
  <cs:wall>
    <cs:lnRef idx="0"/>
    <cs:fillRef idx="0"/>
    <cs:effectRef idx="0"/>
    <cs:fontRef idx="minor">
      <a:schemeClr val="tx1"/>
    </cs:fontRef>
  </cs:wall>
</cs:chartStyle>
	`.trim()
}

function getColors1() {
	return `
	<?xml version="1.0"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">
  <a:schemeClr val="accent1"/>
  <a:schemeClr val="accent2"/>
  <a:schemeClr val="accent3"/>
  <a:schemeClr val="accent4"/>
  <a:schemeClr val="accent5"/>
  <a:schemeClr val="accent6"/>
  <cs:variation/>
  <cs:variation>
    <a:lumMod val="60000"/>
  </cs:variation>
  <cs:variation>
    <a:lumMod val="80000"/>
    <a:lumOff val="20000"/>
  </cs:variation>
  <cs:variation>
    <a:lumMod val="80000"/>
  </cs:variation>
  <cs:variation>
    <a:lumMod val="60000"/>
    <a:lumOff val="40000"/>
  </cs:variation>
  <cs:variation>
    <a:lumMod val="50000"/>
  </cs:variation>
  <cs:variation>
    <a:lumMod val="70000"/>
    <a:lumOff val="30000"/>
  </cs:variation>
  <cs:variation>
    <a:lumMod val="70000"/>
  </cs:variation>
  <cs:variation>
    <a:lumMod val="50000"/>
    <a:lumOff val="50000"/>
  </cs:variation>
</cs:colorStyle>
	`.trim()
}

function Theme2() {

	// @CHRISTOPHER
	// Define Colors
	const colors = ['#4472C4'].map((v: string) => v.replace('#', '').toUpperCase())

	return `
	<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1>
        <a:sysClr val="windowText" lastClr="000000"/>
      </a:dk1>
      <a:lt1>
        <a:sysClr val="window" lastClr="FFFFFF"/>
      </a:lt1>
      <a:dk2>
        <a:srgbClr val="44546A"/>
      </a:dk2>
      <a:lt2>
        <a:srgbClr val="E7E6E6"/>
      </a:lt2>
      <a:accent1>
        <a:srgbClr val="${colors[0]}"/>
      </a:accent1>
      <a:accent2>
        <a:srgbClr val="ED7D31"/>
      </a:accent2>
      <a:accent3>
        <a:srgbClr val="A5A5A5"/>
      </a:accent3>
      <a:accent4>
        <a:srgbClr val="FFC000"/>
      </a:accent4>
      <a:accent5>
        <a:srgbClr val="5B9BD5"/>
      </a:accent5>
      <a:accent6>
        <a:srgbClr val="70AD47"/>
      </a:accent6>
      <a:hlink>
        <a:srgbClr val="0563C1"/>
      </a:hlink>
      <a:folHlink>
        <a:srgbClr val="954F72"/>
      </a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="游ゴシック Light"/>
        <a:font script="Hang" typeface="맑은 고딕"/>
        <a:font script="Hans" typeface="等线 Light"/>
        <a:font script="Hant" typeface="新細明體"/>
        <a:font script="Arab" typeface="Times New Roman"/>
        <a:font script="Hebr" typeface="Times New Roman"/>
        <a:font script="Thai" typeface="Angsana New"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="MoolBoran"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Times New Roman"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
        <a:font script="Geor" typeface="Sylfaen"/>
        <a:font script="Armn" typeface="Arial"/>
        <a:font script="Bugi" typeface="Leelawadee UI"/>
        <a:font script="Bopo" typeface="Microsoft JhengHei"/>
        <a:font script="Java" typeface="Javanese Text"/>
        <a:font script="Lisu" typeface="Segoe UI"/>
        <a:font script="Mymr" typeface="Myanmar Text"/>
        <a:font script="Nkoo" typeface="Ebrima"/>
        <a:font script="Olck" typeface="Nirmala UI"/>
        <a:font script="Osma" typeface="Ebrima"/>
        <a:font script="Phag" typeface="Phagspa"/>
        <a:font script="Syrn" typeface="Estrangelo Edessa"/>
        <a:font script="Syrj" typeface="Estrangelo Edessa"/>
        <a:font script="Syre" typeface="Estrangelo Edessa"/>
        <a:font script="Sora" typeface="Nirmala UI"/>
        <a:font script="Tale" typeface="Microsoft Tai Le"/>
        <a:font script="Talu" typeface="Microsoft New Tai Lue"/>
        <a:font script="Tfng" typeface="Ebrima"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri" panose="020F0502020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="游ゴシック"/>
        <a:font script="Hang" typeface="맑은 고딕"/>
        <a:font script="Hans" typeface="等线"/>
        <a:font script="Hant" typeface="新細明體"/>
        <a:font script="Arab" typeface="Arial"/>
        <a:font script="Hebr" typeface="Arial"/>
        <a:font script="Thai" typeface="Cordia New"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="DaunPenh"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Arial"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
        <a:font script="Geor" typeface="Sylfaen"/>
        <a:font script="Armn" typeface="Arial"/>
        <a:font script="Bugi" typeface="Leelawadee UI"/>
        <a:font script="Bopo" typeface="Microsoft JhengHei"/>
        <a:font script="Java" typeface="Javanese Text"/>
        <a:font script="Lisu" typeface="Segoe UI"/>
        <a:font script="Mymr" typeface="Myanmar Text"/>
        <a:font script="Nkoo" typeface="Ebrima"/>
        <a:font script="Olck" typeface="Nirmala UI"/>
        <a:font script="Osma" typeface="Ebrima"/>
        <a:font script="Phag" typeface="Phagspa"/>
        <a:font script="Syrn" typeface="Estrangelo Edessa"/>
        <a:font script="Syrj" typeface="Estrangelo Edessa"/>
        <a:font script="Syre" typeface="Estrangelo Edessa"/>
        <a:font script="Sora" typeface="Nirmala UI"/>
        <a:font script="Tale" typeface="Microsoft Tai Le"/>
        <a:font script="Talu" typeface="Microsoft New Tai Lue"/>
        <a:font script="Tfng" typeface="Ebrima"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:lumMod val="110000"/>
                <a:satMod val="105000"/>
                <a:tint val="67000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="50000">
              <a:schemeClr val="phClr">
                <a:lumMod val="105000"/>
                <a:satMod val="103000"/>
                <a:tint val="73000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:lumMod val="105000"/>
                <a:satMod val="109000"/>
                <a:tint val="81000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:satMod val="103000"/>
                <a:lumMod val="102000"/>
                <a:tint val="94000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="50000">
              <a:schemeClr val="phClr">
                <a:satMod val="110000"/>
                <a:lumMod val="100000"/>
                <a:shade val="100000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:lumMod val="99000"/>
                <a:satMod val="120000"/>
                <a:shade val="78000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="63000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:solidFill>
          <a:schemeClr val="phClr">
            <a:tint val="95000"/>
            <a:satMod val="170000"/>
          </a:schemeClr>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="93000"/>
                <a:satMod val="150000"/>
                <a:shade val="98000"/>
                <a:lumMod val="102000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="50000">
              <a:schemeClr val="phClr">
                <a:tint val="98000"/>
                <a:satMod val="130000"/>
                <a:shade val="90000"/>
                <a:lumMod val="103000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="63000"/>
                <a:satMod val="120000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
  <a:extLst>
    <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
      <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>
    </a:ext>
  </a:extLst>
</a:theme>
	`.trim()
}
