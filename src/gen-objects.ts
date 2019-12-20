/**
 * PptxGenJS: Slide object generators
 */

import {
	BARCHART_COLORS,
	BASE_SHAPES,
	CHART_TYPE_NAMES,
	CHART_TYPES,
	DEF_CELL_MARGIN_PT,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_SLIDE_MARGIN_IN,
	EMU,
	IMG_PLAYBTN,
	MASTER_OBJECTS,
	ONEPT,
	PIECHART_COLORS,
	SLIDE_OBJECT_TYPES,
	TEXT_HALIGN,
	TEXT_VALIGN,
} from './core-enums'
import {
	IChartMulti,
	IChartOpts,
	IImageOpts,
	ILayout,
	IMediaOpts,
	IShape,
	IShapeOptions,
	ISlide,
	ISlideLayout,
	ISlideMstrObjPlchldrOpts,
	ISlideObject,
	ITableCell,
	ITableOptions,
	IText,
	ITextOpts,
	OptsChartGridLine,
	TableRow,
	ISlideMasterOptions,
} from './core-interfaces'
import { getSlidesForTableRows } from './gen-tables'
import { getSmartParseNumber, inch2Emu, encodeXmlEntities } from './gen-utils'

import TextElement from './elements/text'
import ShapeElement from './elements/simple-shape'
import PlaceholderTextElement from './elements/placeholder-text'
import ImageElement from './elements/image'
import ChartElement from './elements/chart'
import SlideNumberElement from './elements/slide-number'

/** counter for included charts (used for index in their filenames) */
let _chartCounter: number = 0

/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {ISlideMasterOptions} slideDef - slide definition
 * @param {ISlide|ISlideLayout} target - empty slide object that should be updated by the passed definition
 */

/**
 * Adds a media object to a slide definition.
 * @param {ISlide} `target` - slide object that the text will be added to
 * @param {IMediaOpts} `opt` - media options
 */
export function addMediaDefinition(target: ISlide, opt: IMediaOpts) {
	let intRels = target.relsMedia.length + 1
	let intPosX = opt.x || 0
	let intPosY = opt.y || 0
	let intSizeX = opt.w || 2
	let intSizeY = opt.h || 2
	let strData = opt.data || ''
	let strLink = opt.link || ''
	let strPath = opt.path || ''
	let strType = opt.type || 'audio'
	let strExtn = 'mp3'
	let slideData: ISlideObject = {
		type: SLIDE_OBJECT_TYPES.media,
	}

	// STEP 1: REALITY-CHECK
	if (!strPath && !strData && strType !== 'online') {
		throw "addMedia() error: either 'data' or 'path' are required!"
	} else if (strData && strData.toLowerCase().indexOf('base64,') === -1) {
		throw "addMedia() error: `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')"
	}
	// Online Video: requires `link`
	if (strType === 'online' && !strLink) {
		throw 'addMedia() error: online videos require `link` value'
	}

	// FIXME: 20190707
	//strType = strData ? strData.split(';')[0].split('/')[0] : strType
	strExtn = strData ? strData.split(';')[0].split('/')[1] : strPath.split('.').pop()

	// STEP 2: Set type, media
	slideData.mtype = strType
	slideData.media = strPath || 'preencoded.mov'
	slideData.options = {}

	// STEP 3: Set media properties & options
	slideData.options.x = intPosX
	slideData.options.y = intPosY
	slideData.options.w = intSizeX
	slideData.options.h = intSizeY

	// STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
	// NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
	if (strType === 'online') {
		// A: Add video
		target.relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			data: 'dummy',
			type: 'online',
			extn: strExtn,
			rId: intRels + 1,
			Target: strLink,
		})
		slideData.mediaRid = target.relsMedia[target.relsMedia.length - 1].rId

		// B: Add preview/overlay image
		target.relsMedia.push({
			path: 'preencoded.png',
			data: IMG_PLAYBTN,
			type: 'image/png',
			extn: 'png',
			rId: intRels + 2,
			Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.png',
		})
	} else {
		/* NOTE: Audio/Video files consume *TWO* rId's:
		 * <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
		 * <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
		 */

		// A: "relationships/video"
		target.relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: intRels + 0,
			Target: '../media/media-' + target.number + '-' + (target.relsMedia.length + 1) + '.' + strExtn,
		})
		slideData.mediaRid = target.relsMedia[target.relsMedia.length - 1].rId

		// B: "relationships/media"
		target.relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: intRels + 1,
			Target: '../media/media-' + target.number + '-' + (target.relsMedia.length + 0) + '.' + strExtn,
		})

		// C: Add preview/overlay image
		target.relsMedia.push({
			data: IMG_PLAYBTN,
			path: 'preencoded.png',
			type: 'image/png',
			extn: 'png',
			rId: intRels + 2,
			Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.png',
		})
	}

	// LAST
	target.data.push(slideData)
}

/**
 * Adds Notes to a slide.
 * @param {String} `notes`
 * @param {Object} opt (*unused*)
 * @param {ISlide} `target` slide object
 * @since 2.3.0
 */
export function addNotesDefinition(target: ISlide, notes: string) {
	target.data.push({
		type: SLIDE_OBJECT_TYPES.notes,
		text: notes,
	})
}

/**
 * Adds a shape object to a slide definition.
 * @param {IShape} shape shape const object (pptx.shapes)
 * @param {IShapeOptions} opt
 * @param {ISlide} target slide object that the shape should be added to
 */
export function addShapeDefinition(target: ISlide, shape: IShape, opt: IShapeOptions) {
	let options = typeof opt === 'object' ? opt : {}
	let newObject = {
		type: SLIDE_OBJECT_TYPES.text,
		shape: shape,
		options: options,
		text: null,
	}

	// 1: Reality check
	if (!shape || typeof shape !== 'object') throw 'Missing/Invalid shape parameter! Example: `addShape(pptx.shapes.LINE, {x:1, y:1, w:1, h:1});`'

	// 2: Set options defaults
	options.x = options.x || (options.x === 0 ? 0 : 1)
	options.y = options.y || (options.y === 0 ? 0 : 1)
	options.w = options.w || (options.w === 0 ? 0 : 1)
	options.h = options.h || (options.h === 0 ? 0 : 1)
	options.line = options.line || (shape.name === 'line' ? '333333' : null)
	options.lineSize = options.lineSize || (shape.name === 'line' ? 1 : null)
	if (['dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'solid', 'sysDash', 'sysDot'].indexOf(options.lineDash || '') < 0) options.lineDash = 'solid'

	// 3: Add object to slide
	target.data.push(newObject)
}

/**
 * Adds a table object to a slide definition.
 * @param {ISlide} target - slide object that the table should be added to
 * @param {TableRow[]} arrTabRows - table data
 * @param {ITableOptions} inOpt - table options
 * @param {ISlideLayout} slideLayout - Slide layout
 * @param {ILayout} presLayout - Presenation layout
 * @param {Function} addSlide - method
 * @param {Function} getSlide - method
 */
export function addTableDefinition(
	target: ISlide,
	tableRows: TableRow[],
	options: ITableOptions,
	slideLayout: ISlideLayout,
	presLayout: ILayout,
	addSlide: Function,
	getSlide: Function
) {
	let opt: ITableOptions = options && typeof options === 'object' ? options : {}
	let slides: ISlide[] = [target] // Create array of Slides as more may be added by auto-paging

	// STEP 1: REALITY-CHECK
	{
		// A: check for empty
		if (tableRows === null || tableRows.length === 0 || !Array.isArray(tableRows)) {
			throw `addTable: Array expected! EX: 'slide.addTable( [rows], {options} );' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)`
		}

		// B: check for non-well-formatted array (ex: rows=['a','b'] instead of [['a','b']])
		if (!tableRows[0] || !Array.isArray(tableRows[0])) {
			throw `addTable: 'rows' should be an array of cells! EX: 'slide.addTable( [ ['A'], ['B'], {text:'C',options:{align:'center'}} ] );' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)`
		}
	}

	// STEP 2: Transform `tableRows` into well-formatted ITableCell's
	// tableRows can be object or plain text array: `[{text:'cell 1'}, {text:'cell 2', options:{color:'ff0000'}}]` | `["cell 1", "cell 2"]`
	let arrRows: [ITableCell[]?] = []
	tableRows.forEach(row => {
		let newRow: ITableCell[] = []

		if (Array.isArray(row)) {
			row.forEach((cell: number | string | ITableCell) => {
				let newCell: ITableCell = {
					type: SLIDE_OBJECT_TYPES.tablecell,
					text: '',
					options: typeof cell === 'object' ? cell.options : null,
				}
				if (typeof cell === 'string' || typeof cell === 'number') newCell.text = cell.toString()
				else if (cell.text) {
					// Cell can contain complex text type, or string, or number
					if (typeof cell.text === 'string' || typeof cell.text === 'number') newCell.text = cell.text.toString()
					else if (cell.text) newCell.text = cell.text
					// Capture options
					if (cell.options) newCell.options = cell.options
				}
				newRow.push(newCell)
			})
		} else {
			console.log('addTable: tableRows has a bad row. A row should be an array of cells. You provided:')
			console.log(row)
		}

		arrRows.push(newRow)
	})

	// STEP 3: Set options
	opt.x = getSmartParseNumber(opt.x || (opt.x === 0 ? 0 : EMU / 2), 'X', presLayout)
	opt.y = getSmartParseNumber(opt.y || (opt.y === 0 ? 0 : EMU / 2), 'Y', presLayout)
	if (opt.h) opt.h = getSmartParseNumber(opt.h, 'Y', presLayout) // NOTE: Dont set default `h` - leaving it null triggers auto-rowH in `makeXMLSlide()`
	opt.autoPage = typeof opt.autoPage === 'boolean' ? opt.autoPage : false
	opt.fontSize = opt.fontSize || DEF_FONT_SIZE
	opt.autoPageLineWeight = typeof opt.autoPageLineWeight !== 'undefined' && !isNaN(Number(opt.autoPageLineWeight)) ? Number(opt.autoPageLineWeight) : 0
	opt.margin = opt.margin === 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_PT
	if (typeof opt.margin === 'number') opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)]
	if (opt.autoPageLineWeight > 1) opt.autoPageLineWeight = 1
	else if (opt.autoPageLineWeight < -1) opt.autoPageLineWeight = -1
	// Set default color if needed (table option > inherit from Slide > default to black)
	if (!opt.color) opt.color = opt.color || DEF_FONT_COLOR

	// Set/Calc table width
	// Get slide margins - start with default values, then adjust if master or slide margins exist
	let arrTableMargin = DEF_SLIDE_MARGIN_IN
	// Case 1: Master margins
	if (slideLayout && typeof slideLayout.margin !== 'undefined') {
		if (Array.isArray(slideLayout.margin)) arrTableMargin = slideLayout.margin
		else if (!isNaN(Number(slideLayout.margin)))
			arrTableMargin = [Number(slideLayout.margin), Number(slideLayout.margin), Number(slideLayout.margin), Number(slideLayout.margin)]
	}
	// Case 2: Table margins
	/* FIXME: add `margin` option to slide options
		else if ( addNewSlide.margin ) {
			if ( Array.isArray(addNewSlide.margin) ) arrTableMargin = addNewSlide.margin;
			else if ( !isNaN(Number(addNewSlide.margin)) ) arrTableMargin = [Number(addNewSlide.margin), Number(addNewSlide.margin), Number(addNewSlide.margin), Number(addNewSlide.margin)];
		}
	*/

	// Calc table width depending upon what data we have - several scenarios exist (including bad data, eg: colW doesnt match col count)
	if (opt.w) {
		opt.w = getSmartParseNumber(opt.w, 'X', presLayout)
	} else if (opt.colW) {
		if (typeof opt.colW === 'string' || typeof opt.colW === 'number') {
			opt.w = Math.floor(Number(opt.colW) * arrRows[0].length)
		} else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length !== arrRows[0].length) {
			console.warn('addTable: colW.length != data.length! Defaulting to evenly distributed col widths.')

			let numColWidth = Math.floor((presLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3]) / arrRows[0].length)
			opt.colW = []
			for (let idx = 0; idx < arrRows[0].length; idx++) {
				opt.colW.push(numColWidth)
			}
			opt.w = Math.floor(numColWidth * arrRows[0].length)
		}
	} else {
		opt.w = Math.floor(presLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3])
	}

	// STEP 4: Convert units to EMU now (we use different logic in makeSlide->table - smartCalc is not used)
	if (opt.x && opt.x < 20) opt.x = inch2Emu(opt.x)
	if (opt.y && opt.y < 20) opt.y = inch2Emu(opt.y)
	if (opt.w && opt.w < 20) opt.w = inch2Emu(opt.w)
	if (opt.h && opt.h < 20) opt.h = inch2Emu(opt.h)

	// STEP 5: Loop over cells: transform each to ITableCell; check to see whether to skip autopaging while here
	arrRows.forEach(row => {
		row.forEach((cell, idy) => {
			// A: Transform cell data if needed
			/* Table rows can be an object or plain text - transform into object when needed
				// EX:
				var arrTabRows1 = [
					[ { text:'A1\nA2', options:{rowspan:2, fill:'99FFCC'} } ]
					,[ 'B2', 'C2', 'D2', 'E2' ]
				]
			*/
			if (typeof cell === 'number' || typeof cell === 'string') {
				// Grab table formatting `opts` to use here so text style/format inherits as it should
				row[idy] = { type: SLIDE_OBJECT_TYPES.tablecell, text: row[idy].toString(), options: opt }
			} else if (typeof cell === 'object') {
				// ARG0: `text`
				if (typeof cell.text === 'number') row[idy].text = row[idy].text.toString()
				else if (typeof cell.text === 'undefined' || cell.text === null) row[idy].text = ''

				// ARG1: `options`: ensure options exists
				row[idy].options = cell.options || {}

				// Set type to tabelcell
				row[idy].type = SLIDE_OBJECT_TYPES.tablecell
			}

			// B: Check for fine-grained formatting, disable auto-page when found
			// Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
			// Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
			if (cell.text && Array.isArray(cell.text)) opt.autoPage = false
		})
	})

	// STEP 6: Auto-Paging: (via {options} and used internally)
	// (used internally by `tableToSlides()` to not engage recursion - we've already paged the table data, just add this one)
	if (opt && opt.autoPage === false) {
		// Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
		createHyperlinkRels(target, arrRows)

		// Add data (NOTE: Use `extend` to avoid mutation)
		target.data.push({
			type: SLIDE_OBJECT_TYPES.table,
			arrTabRows: arrRows,
			options: Object.assign({}, opt),
		})
	} else {
		// Loop over rows and create 1-N tables as needed (ISSUE#21)
		getSlidesForTableRows(arrRows, opt, presLayout, slideLayout).forEach((slide, idx) => {
			// A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
			if (!getSlide(target.number + idx)) slides.push(addSlide(slideLayout ? slideLayout.name : null))

			// B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
			if (idx > 0) opt.y = inch2Emu(opt.newSlideStartY || arrTableMargin[0])

			// C: Add this table to new Slide
			{
				let newSlide: ISlide = getSlide(target.number + idx)

				opt.autoPage = false

				// Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
				createHyperlinkRels(newSlide, slide.rows)

				// Add rows to new slide
				newSlide.addTable(slide.rows, Object.assign({}, opt))
			}
		})
	}
}

/**
 * Adds placeholder objects to slide
 * @param {ISlide} slide - slide object containing layouts
 */
export function addPlaceholdersToSlideLayouts(slide: ISlide) {
	// Add all placeholders on this Slide that dont already exist
	;(slide.slideLayout.data || []).forEach(slideLayoutObj => {
		if (slideLayoutObj instanceof PlaceholderTextElement) {
			// A: Search for this placeholder on Slide before we add
			// NOTE: Check to ensure a placeholder does not already exist on the Slide
			// They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
			if (
				slide.data.filter(slideObj => {
					const placeholder = slideObj.placeholder || (slideObj.options && slideObj.options.placeholder)
					return placeholder === slideLayoutObj.name
				}).length === 0
			) {
				if (slideLayoutObj.placeholderType !== 'pic') {
					slide.data.push(new TextElement('', { placeholder: slideLayoutObj.name }, () => null))
				}
			}
		}
	})
}

/* -------------------------------------------------------------------------------- */

/**
 * Adds a background image or color to a slide definition.
 * @param {String|Object} bkg - color string or an object with image definition
 * @param {ISlide} target - slide object that the background is set to
 */
function addBackgroundDefinition(bkg: string | { src?: string; path?: string; data?: string }, target: ISlide | ISlideLayout) {
	if (typeof bkg === 'object' && (bkg.src || bkg.path || bkg.data)) {
		// Allow the use of only the data key (`path` isnt reqd)
		bkg.src = bkg.src || bkg.path || null
		if (!bkg.src) bkg.src = 'preencoded.png'
		let strImgExtn = (bkg.src.split('.').pop() || 'png').split('?')[0] // Handle "blah.jpg?width=540" etc.
		if (strImgExtn === 'jpg') strImgExtn = 'jpeg' // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup

		let intRels = target.relsMedia.length + 1
		target.relsMedia.push({
			path: bkg.src,
			type: SLIDE_OBJECT_TYPES.image,
			extn: strImgExtn,
			data: bkg.data || null,
			rId: intRels,
			Target: '../media/image' + (target.relsMedia.length + 1) + '.' + strImgExtn,
		})
		target.bkgdImgRid = intRels
	} else if (bkg && typeof bkg === 'string') {
		target.bkgd = bkg
	}
}

/**
 * Parses text/text-objects from `addText()` and `addTable()` methods; creates 'hyperlink'-type Slide Rels for each hyperlink found
 * @param {ISlide} target - slide object that any hyperlinks will be be added to
 * @param {number | string | IText | IText[] | ITableCell[][]} text - text to parse
 */
function createHyperlinkRels(target: ISlide, text: number | string | IText | IText[] | ITableCell[][]) {
	let textObjs = []

	// Only text objects can have hyperlinks, bail when text param is plain text
	if (typeof text === 'string' || typeof text === 'number') return
	// IMPORTANT: "else if" Array.isArray must come before typeof===object! Otherwise, code will exhaust recursion!
	else if (Array.isArray(text)) textObjs = text
	else if (typeof text === 'object') textObjs = [text]

	textObjs.forEach((text: IText) => {
		// `text` can be an array of other `text` objects (table cell word-level formatting), continue parsing using recursion
		if (Array.isArray(text)) createHyperlinkRels(target, text)
		else if (text && typeof text === 'object' && text.options && text.options.hyperlink && !text.options.hyperlink.rId) {
			if (typeof text.options.hyperlink !== 'object') console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:'https://github.com'}` ")
			else if (!text.options.hyperlink.url && !text.options.hyperlink.slide) console.log("ERROR: 'hyperlink requires either: `url` or `slide`'")
			else {
				let relId = target.rels.length + target.relsChart.length + target.relsMedia.length + 1

				target.rels.push({
					type: SLIDE_OBJECT_TYPES.hyperlink,
					data: text.options.hyperlink.slide ? 'slide' : 'dummy',
					rId: relId,
					Target: encodeXmlEntities(text.options.hyperlink.url) || text.options.hyperlink.slide.toString(),
				})

				text.options.hyperlink.rId = relId
			}
		}
	})
}
