/**
 * PptxGenJS: Slide object generators
 */

import {
	CRLF,
	EMU,
	ONEPT,
	MASTER_OBJECTS,
	BARCHART_COLORS,
	PIECHART_COLORS,
	DEF_CELL_BORDER,
	DEF_CELL_MARGIN_PT,
	CHART_TYPES,
	SLDNUMFLDID,
	BULLET_TYPES,
	DEF_FONT_COLOR,
	LAYOUT_IDX_SERIES_BASE,
	PLACEHOLDER_TYPES,
	SLIDE_OBJECT_TYPES,
	DEF_FONT_SIZE,
	DEF_SLIDE_MARGIN_IN,
	IMG_PLAYBTN,
	BASE_SHAPES,
    CHART_TYPE_NAMES,
} from './enums'
import { ISlide, ITextOpts, ILayout, ISlideLayout, ISlideDataObject, ITableCell, ISlideLayoutData, IMediaOpts, ISlideRelMedia, IChartOpts, IChartMulti } from './interfaces'
import { convertRotationDegrees, encodeXmlEntities, getSmartParseNumber, inch2Emu, genXmlColorSelection } from './utils'
import { createHyperlinkRels, getSlidesForTableRows, correctShadowOptions, correctGridLineOptions, getShapeInfo, genXmlTextBody, genXmlPlaceholder } from './gen-xml'

/** counter for included images (used for index in their filenames) */
var _imageCounter: number = 0
/** counter for included charts (used for index in their filenames) */
var _chartCounter: number = 0

export function addTableDefinition(target: ISlide, arrTabRows, inOpt, slideLayout: ISlideLayout, presLayout: ILayout) {
	var opt = inOpt && typeof inOpt === 'object' ? inOpt : {}

	// STEP 1: REALITY-CHECK
	if (arrTabRows == null || arrTabRows.length == 0 || !Array.isArray(arrTabRows)) {
		try {
			console.warn('[warn] addTable: Array expected! USAGE: slide.addTable( [rows], {options} );')
		} catch (ex) {}
		return null
	}

	// STEP 2: Row setup: Handle case where user passed in a simple 1-row array. EX: `["cell 1", "cell 2"]`
	//var arrRows = jQuery.extend(true,[],arrTabRows);
	//if ( !Array.isArray(arrRows[0]) ) arrRows = [ jQuery.extend(true,[],arrTabRows) ];
	var arrRows = arrTabRows
	if (!Array.isArray(arrRows[0])) arrRows = [arrTabRows]

	// STEP 3: Set options
	opt.x = getSmartParseNumber(opt.x || (opt.x == 0 ? 0 : EMU / 2), 'X', slideLayout)
	opt.y = getSmartParseNumber(opt.y || (opt.y == 0 ? 0 : EMU), 'Y', slideLayout)
	opt.cy = opt.h || opt.cy // NOTE: Dont set default `cy` - leaving it null triggers auto-rowH in `makeXMLSlide()`
	if (opt.cy) opt.cy = getSmartParseNumber(opt.cy, 'Y', slideLayout)
	opt.h = opt.cy
	opt.autoPage = opt.autoPage == false ? false : true
	opt.fontSize = opt.fontSize || DEF_FONT_SIZE
	opt.lineWeight = typeof opt.lineWeight !== 'undefined' && !isNaN(Number(opt.lineWeight)) ? Number(opt.lineWeight) : 0
	opt.margin = opt.margin == 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_PT
	if (!isNaN(opt.margin)) opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)]
	if (opt.lineWeight > 1) opt.lineWeight = 1
	else if (opt.lineWeight < -1) opt.lineWeight = -1
	// Set default color if needed (table option > inherit from Slide > default to black)
	if (!opt.color) opt.color = opt.color || DEF_FONT_COLOR

	// Set/Calc table width
	// Get slide margins - start with default values, then adjust if master or slide margins exist
	var arrTableMargin = DEF_SLIDE_MARGIN_IN
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
	if (opt.w || opt.cx) {
		opt.cx = getSmartParseNumber(opt.w || opt.cx, 'X', slideLayout)
		opt.w = opt.cx
	} else if (opt.colW) {
		if (typeof opt.colW === 'string' || typeof opt.colW === 'number') {
			opt.cx = Math.floor(Number(opt.colW) * arrRows[0].length)
			opt.w = opt.cx
		} else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length != arrRows[0].length) {
			console.warn('addTable: colW.length != data.length! Defaulting to evenly distributed col widths.')

			var numColWidth = Math.floor((presLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3]) / arrRows[0].length)
			opt.colW = []
			for (var idx = 0; idx < arrRows[0].length; idx++) {
				opt.colW.push(numColWidth)
			}
			opt.cx = Math.floor(numColWidth * arrRows[0].length)
			opt.w = opt.cx
		}
	} else {
		var numTabWidth = presLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3]
		opt.cx = Math.floor(numTabWidth)
		opt.w = opt.cx
	}

	// STEP 4: Convert units to EMU now (we use different logic in makeSlide->table - smartCalc is not used)
	if (opt.x < 20) opt.x = inch2Emu(opt.x)
	if (opt.y < 20) opt.y = inch2Emu(opt.y)
	if (opt.cx < 20) opt.cx = inch2Emu(opt.cx)
	if (opt.cy && opt.cy < 20) opt.cy = inch2Emu(opt.cy)

	// STEP 5: Check for fine-grained formatting, disable auto-page when found
	// Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
	// Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
	arrRows.forEach(row => {
		row.forEach(cell => {
			if (cell && cell.text && Array.isArray(cell.text)) opt.autoPage = false
		})
	})

	// STEP 6: Create hyperlink rels
	createHyperlinkRels(this.slides, arrRows, this.rels)
	// FIXME: why do we need all slides???

	// STEP 7: Auto-Paging: (via {options} and used internally)
	// (used internally by `addSlidesForTable()` to not engage recursion - we've already paged the table data, just add this one)
	if (opt && opt.autoPage == false) {
		// Add data (NOTE: Use `extend` to avoid mutation)
		//gObjPptx.slides[slideNum].data[gObjPptx.slides[slideNum].data.length] = {
		// FIXME: create an addObject method instead of touching `data`!!
		target.data.push({
			type: SLIDE_OBJECT_TYPES.table,
			arrTabRows: arrRows,
			options: jQuery.extend(true, {}, opt),
		})
	} else {
		// Loop over rows and create 1-N tables as needed (ISSUE#21)
		getSlidesForTableRows(arrRows, opt, presLayout).forEach((arrRows, idx) => {
			// A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
			//let currSlide = !this.slides[slideNum + idx] ? this.addNewSlide(inMasterName) : this.slides[slideNum + idx]
			// FIXME: ^^^

			// B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
			if (idx > 0) opt.y = inch2Emu(opt.newPageStartY || arrTableMargin[0])

			// C: Add this table to new Slide
			opt.autoPage = false
			target.addTable(arrRows, jQuery.extend(true, {}, opt))
		})
	}
}

/**
 * Adds a background image or color to a slide definition.
 * @param {String|Object} bkg color string or an object with image definition
 * @param {ISlide} target slide object that the background is set to
 */
export function addBackgroundDefinition(bkg: string | { src?: string; path?: string; data?: string }, target: ISlide) {
	if (typeof bkg === 'object' && (bkg.src || bkg.path || bkg.data)) {
		// Allow the use of only the data key (`path` isnt reqd)
		bkg.src = bkg.src || bkg.path || null
		if (!bkg.src) bkg.src = 'preencoded.png'
		var targetRels = target.rels
		var strImgExtn = (bkg.src.split('.').pop() || 'png').split('?')[0] // Handle "blah.jpg?width=540" etc.
		if (strImgExtn == 'jpg') strImgExtn = 'jpeg' // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup

		var intRels = targetRels.length + 1
		targetRels.push({
			path: bkg.src,
			type: SLIDE_OBJECT_TYPES.image,
			extn: strImgExtn,
			data: bkg.data || null,
			rId: intRels,
			Target: '../media/image' + ++_imageCounter + '.' + strImgExtn,
		})
		target.bkgdImgRid = intRels
	} else if (bkg && typeof bkg === 'string') {
		target.bkgd = bkg
	}
}

/**
 * Adds a text object to a slide definition.
 * @param {String} text
 * @param {ITextOpts} opt
 * @param {ISlide} target - slide object that the text should be added to
 * @param {Boolean} isPlaceholder
 * @since: 1.0.0
 */
export function addTextDefinition(text: string | Array<object>, opt: ITextOpts, target: ISlide, isPlaceholder: boolean) {
	var opt: ITextOpts = opt || {}
	var text = text || ''
	if (Array.isArray(text) && text.length == 0) text = ''
	var resultObject = {
		type: null,
		text: null,
		options: null,
	}

	// STEP 2: Set some options
	// Placeholders should inherit their colors or override them, so don't default them
	if (!opt.placeholder) {
		opt.color = opt.color || target.color || DEF_FONT_COLOR // Set color (options > inherit from Slide > default to black)
	}

	// ROBUST: Convert attr values that will likely be passed by users to valid OOXML values
	if (opt.valign)
		opt.valign = opt.valign
			.toLowerCase()
			.replace(/^c.*/i, 'ctr')
			.replace(/^m.*/i, 'ctr')
			.replace(/^t.*/i, 't')
			.replace(/^b.*/i, 'b')
	if (opt.align)
		opt.align = opt.align
			.toLowerCase()
			.replace(/^c.*/i, 'center')
			.replace(/^m.*/i, 'center')
			.replace(/^l.*/i, 'left')
			.replace(/^r.*/i, 'right')

	// ROBUST: Set rational values for some shadow props if needed
	correctShadowOptions(opt.shadow)

	// STEP 3: Set props
	resultObject.type = isPlaceholder ? 'placeholder' : 'text'
	resultObject.text = text

	// STEP 4: Set options
	resultObject.options = opt
	if (opt.shape && opt.shape.name == 'line') {
		opt.line = opt.line || '333333'
		opt.lineSize = opt.lineSize || 1
	}
	resultObject.options.bodyProp = {}
	resultObject.options.bodyProp.autoFit = opt.autoFit || false // If true, shape will collapse to text size (Fit To Shape)
	resultObject.options.bodyProp.anchor = opt.valign || (!opt.placeholder ? 'ctr' : null) // VALS: [t,ctr,b]
	resultObject.options.bodyProp.rot = opt.rotate || null // VALS: degree * 60,000
	resultObject.options.bodyProp.vert = opt.vert || null // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
	resultObject.options.lineSpacing = opt.lineSpacing && !isNaN(opt.lineSpacing) ? opt.lineSpacing : null

	if ((opt.inset && !isNaN(Number(opt.inset))) || opt.inset == 0) {
		resultObject.options.bodyProp.lIns = inch2Emu(opt.inset)
		resultObject.options.bodyProp.rIns = inch2Emu(opt.inset)
		resultObject.options.bodyProp.tIns = inch2Emu(opt.inset)
		resultObject.options.bodyProp.bIns = inch2Emu(opt.inset)
	}

	target.data.push(resultObject)
	createHyperlinkRels([target], text || '', target.rels)

	return resultObject
}

/**
 * Adds a media object to a slide definition.
 * @param {ISlide} target - slide object that the text should be added to
 * @param {IMediaOpts} opt
 */
export function addMediaDefinition(target: ISlide, opt: IMediaOpts) {
	let intRels = 1
	let intImages = ++this._imageCounter // FIXME: _imageCounter doesnt exist here!
	let intPosX = opt.x || 0
	let intPosY = opt.y || 0
	let intSizeX = opt.w || 2
	let intSizeY = opt.h || 2
	let strData = opt.data || ''
	let strLink = opt.link || ''
	let strPath = opt.path || ''
	let strType = opt.type || 'audio'
	let strExtn = 'mp3'

	// STEP 1: REALITY-CHECK
	if (!strPath && !strData && strType != 'online') {
		throw "addMedia() error: either 'data' or 'path' are required!"
	} else if (strData && strData.toLowerCase().indexOf('base64,') == -1) {
		throw "addMedia() error: `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')"
	}
	// Online Video: requires `link`
	if (strType == 'online' && !strLink) {
		throw 'addMedia() error: online videos require `link` value'
	}

	// FIXME: 20190707
	//strType = strData ? strData.split(';')[0].split('/')[0] : strType
	strExtn = strData ? strData.split(';')[0].split('/')[1] : strPath.split('.').pop()

	// TODO: FIXME: use method instead of pushing into data!
	let slideData = {
		type: SLIDE_OBJECT_TYPES.media,
		mtype: strType,
		media: strPath || 'preencoded.mov',
		options: {},
	} as ISlideDataObject

	// STEP 3: Set media properties & options
	slideData.options.x = intPosX
	slideData.options.y = intPosY
	slideData.options.cx = intSizeX
	slideData.options.cy = intSizeY

	// STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
	// NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
	this.slides.forEach(slide => {
		intRels += slide.rels.length
	})
	// FIXME: "this.slides" doesnt exist - pass in or create method!

	if (strType == 'online') {
		// Add video
		target.relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			data: 'dummy',
			type: 'online',
			extn: strExtn,
			rId: intRels + 1,
			Target: strLink,
		})
		slideData.mediaRid = target.relsMedia[target.relsMedia.length - 1].rId

		// Add preview/overlay image
		target.relsMedia.push({
			path: 'preencoded.png',
			data: IMG_PLAYBTN,
			type: 'image/png',
			extn: 'png',
			rId: intRels + 2,
			Target: '../media/image' + intRels + '.png',
		})
	} else {
		let objRel: ISlideRelMedia = {
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: intRels + 0,
			Target: '../media/media' + intImages + '.' + strExtn,
		}
		// Audio/Video files consume *TWO* rId's:
		// <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
		// <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
		target.relsMedia.push(objRel)
		slideData.mediaRid = target.relsMedia[target.relsMedia.length - 1].rId
		target.relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: intRels + 1,
			Target: '../media/media' + intImages + '.' + strExtn,
		})
		// Add preview/overlay image
		target.relsMedia.push({
			data: IMG_PLAYBTN,
			path: 'preencoded.png',
			type: 'image/png',
			extn: 'png',
			rId: intRels + 2,
			Target: '../media/image' + intImages + '.png',
		})
	}

	target.data.push(slideData)
}

/**
 * Adds Notes to a slide.
 * @param {String} `notes`
 * @param {Object} opt (*unused*)
 * @param {ISlide} `target` slide object
 * @since 2.3.0
 */
export function addNotesDefinition(notes: string, opt: object, target: ISlide) {
	var opt = opt && typeof opt === 'object' ? opt : {}
	var resultObject: ISlideDataObject = {
		type: null,
		text: null,
	}

	resultObject.type = SLIDE_OBJECT_TYPES.notes
	resultObject.text = notes

	target.data.push(resultObject)

	return resultObject
}

/**
 * Adds a placeholder object to a slide definition.
 * @param {String} `text`
 * @param {Object} `opt`
 * @param {ISlide} `target` slide object that the placeholder should be added to
 */
export function addPlaceholderDefinition(text: string, opt: object, target: ISlide) {
	return addTextDefinition(text, opt, target, true)
}

/**
 * Adds a shape object to a slide definition.
 * @param {gObjPptxShapes} shape shape const object (pptx.shapes)
 * @param {Object} opt
 * @param {Object} target slide object that the shape should be added to
 * @return {Object} shape object
 */
export function addShapeDefinition(shape, opt, target) {
	var options = typeof opt === 'object' ? opt : {}
	var resultObject = {
		type: null,
		text: null,
		options: {},
	}

	if (!shape || typeof shape !== 'object') {
		console.error('Missing/Invalid shape parameter! Example: `addShape(pptx.shapes.LINE, {x:1, y:1, w:1, h:1});` ')
		return
	}

	resultObject.type = 'text'
	resultObject.options = options
	options.shape = shape
	options.x = options.x || (options.x == 0 ? 0 : 1)
	options.y = options.y || (options.y == 0 ? 0 : 1)
	options.w = options.w || (options.w == 0 ? 0 : 1)
	options.h = options.h || (options.h == 0 ? 0 : 1)
	options.line = options.line || (shape.name == 'line' ? '333333' : null)
	options.lineSize = options.lineSize || (shape.name == 'line' ? 1 : null)
	if (['dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'solid', 'sysDash', 'sysDot'].indexOf(options.lineDash || '') < 0) options.lineDash = 'solid'

	target.data.push(resultObject)
	return resultObject
}

/**
 * Adds an image object to a slide definition.
 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
 * @param {Object} objImage - object containing `path`/`data`, `x`, `y`, etc.
 * @param {Object} target - slide that the image should be added to (if not specified as the 2nd arg)
 * @return {Object} image object
 */
export function addImageDefinition(objImage, target) {
	var resultObject = {
		type: null,
		text: null,
		options: null,
		image: null,
		imageRid: null,
		hyperlink: null,
	}
	// FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
	var intPosX = objImage.x || 0
	var intPosY = objImage.y || 0
	var intWidth = objImage.w || 0
	var intHeight = objImage.h || 0
	var sizing = objImage.sizing || null
	var objHyperlink = objImage.hyperlink || ''
	var strImageData = objImage.data || ''
	var strImagePath = objImage.path || ''
	var imageRelId = target.rels.length + 1

	// REALITY-CHECK:
	if (!strImagePath && !strImageData) {
		console.error("ERROR: `addImage()` requires either 'data' or 'path' parameter!")
		return null
	} else if (strImageData && strImageData.toLowerCase().indexOf('base64,') == -1) {
		console.error("ERROR: Image `data` value lacks a base64 header! Ex: 'image/png;base64,NMP[...]')")
		return null
	}

	// STEP 1: Set extension
	// NOTE: Split to address URLs with params (eg: `path/brent.jpg?someParam=true`)
	var strImgExtn =
		strImagePath
			.split('.')
			.pop()
			.split('?')[0]
			.split('#')[0] || 'png'
	// However, pre-encoded images can be whatever mime-type they want (and good for them!)
	if (strImageData && /image\/(\w+)\;/.exec(strImageData) && /image\/(\w+)\;/.exec(strImageData).length > 0) {
		strImgExtn = /image\/(\w+)\;/.exec(strImageData)[1]
	} else if (strImageData && strImageData.toLowerCase().indexOf('image/svg+xml') > -1) {
		strImgExtn = 'svg'
	}
	// STEP 2: Set type/path
	resultObject.type = 'image'
	resultObject.image = strImagePath || 'preencoded.png'

	// STEP 3: Set image properties & options
	// FIXME: Measure actual image when no intWidth/intHeight params passed
	// ....: This is an async process: we need to make getSizeFromImage use callback, then set H/W...
	// if ( !intWidth || !intHeight ) { var imgObj = getSizeFromImage(strImagePath);
	var imgObj = { width: 1, height: 1 }
	resultObject.options = {
		x: intPosX || 0,
		y: intPosY || 0,
		cx: intWidth || imgObj.width,
		cy: intHeight || imgObj.height,
		rounding: objImage.rounding || false,
		sizing: sizing,
		placeholder: objImage.placeholder,
	}

	// STEP 4: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	if (strImgExtn == 'svg') {
		// SVG files consume *TWO* rId's: (a png version and the svg image)
		// <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
		// <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>

		target.rels.push({
			path: strImagePath || strImageData + 'png',
			type: 'image/png',
			extn: 'png',
			data: strImageData || '',
			rId: imageRelId,
			Target: '../media/image' + ++_imageCounter + '.png',
			isSvgPng: true,
			svgSize: { w: resultObject.options.cx, h: resultObject.options.cy },
		})
		resultObject.imageRid = imageRelId
		target.rels.push({
			path: strImagePath || strImageData,
			type: 'image/' + strImgExtn,
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId + 1,
			Target: '../media/image' + ++_imageCounter + '.' + strImgExtn,
		})
		resultObject.imageRid = imageRelId + 1
	} else {
		target.rels.push({
			path: strImagePath || 'preencoded.' + strImgExtn,
			type: 'image/' + strImgExtn,
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId,
			Target: '../media/image' + ++_imageCounter + '.' + strImgExtn,
		})
		resultObject.imageRid = imageRelId
	}

	// STEP 5: (Issue#77) Hyperlink support
	if (typeof objHyperlink === 'object') {
		if (!objHyperlink.url && !objHyperlink.slide) console.log("ERROR: 'hyperlink requires either: `url` or `slide`'")
		else {
			var intRelId = imageRelId + 1

			target.rels.push({
				type: 'hyperlink',
				data: objHyperlink.slide ? 'slide' : 'dummy',
				rId: intRelId,
				Target: objHyperlink.url || objHyperlink.slide,
			})

			objHyperlink.rId = intRelId
			resultObject.hyperlink = objHyperlink
		}
	}

	target.data.push(resultObject)
	return resultObject
}

/**
 * Generate the chart based on input data.
 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
 *
 * @param {object} type should belong to: 'column', 'pie'
 * @param {object} data a JSON object with follow the following format
 * @param {object} opt
 * @param {object} target slide object that the chart should be added to
 * @return {Object} chart object
 * {
 *   title: 'eSurvey chart',
 *   data: [
 *		{
 *			name: 'Income',
 *			labels: ['2005', '2006', '2007', '2008', '2009'],
 *			values: [23.5, 26.2, 30.1, 29.5, 24.6]
 *		},
 *		{
 *			name: 'Expense',
 *			labels: ['2005', '2006', '2007', '2008', '2009'],
 *			values: [18.1, 22.8, 23.9, 25.1, 25]
 *		}
 *	 ]
 *	}
 */
export function addChartDefinition(type:CHART_TYPE_NAMES|IChartMulti[], data:[], opt, target:ISlide) {
	var targetRels = target.rels
	var chartId = ++_chartCounter
	var chartRelId = target.rels.length + 1
	var resultObject = {
		type: null,
		text: null,
		options: null,
		chartRid: null,
	}
	// DESIGN: `type` can an object (ex: `pptx.charts.DOUGHNUT`) or an array of chart objects
	// EX: addChartDefinition([ { type:pptx.charts.BAR, data:{name:'', labels:[], values[]} }, {<etc>} ])
	// Multi-Type Charts
	var tmpOpt
	var tmpData = [],
		options:IChartOpts
	if (Array.isArray(type)) {
		// For multi-type charts there needs to be data for each type,
		// as well as a single data source for non-series operations.
		// The data is indexed below to keep the data in order when segmented
		// into types.
		type.forEach(obj => {
			tmpData = tmpData.concat(obj.data)
		})
		tmpOpt = data || opt
	} else {
		tmpData = data
		tmpOpt = opt
	}
	tmpData.forEach((item, i) => {
		item.index = i
	})
	options = tmpOpt && typeof tmpOpt === 'object' ? tmpOpt : {}

	// STEP 1: TODO: check for reqd fields, correct type, etc
	// `type` exists in CHART_TYPES
	// Array.isArray(data)
	/*
		if ( Array.isArray(rel.data) && rel.data.length > 0 && typeof rel.data[0] === 'object'
			&& rel.data[0].labels && Array.isArray(rel.data[0].labels)
			&& rel.data[0].values && Array.isArray(rel.data[0].values) ) {
			obj = rel.data[0];
		}
		else {
			console.warn("USAGE: addChart( 'pie', [ {name:'Sales', labels:['Jan','Feb'], values:[10,20]} ], {x:1, y:1} )");
			return;
		}
		*/

	// STEP 2: Set default options/decode user options
	// A: Core
	options.type = type
	options.x = typeof options.x !== 'undefined' && options.x != null && !isNaN(Number(options.x)) ? options.x : 1
	options.y = typeof options.y !== 'undefined' && options.y != null && !isNaN(Number(options.y)) ? options.y : 1
	options.w = options.w || '50%'
	options.h = options.h || '50%'

	// B: Options: misc
	if (['bar', 'col'].indexOf(options.barDir || '') < 0) options.barDir = 'col'
	// IMPORTANT: 'bestFit' will cause issues with PPT-Online in some cases, so defualt to 'ctr'!
	if (['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't'].indexOf(options.dataLabelPosition || '') < 0)
		options.dataLabelPosition = options.type == CHART_TYPES.PIE || options.type == CHART_TYPES.DOUGHNUT ? 'bestFit' : 'ctr'
	options.dataLabelBkgrdColors = options.dataLabelBkgrdColors == true || options.dataLabelBkgrdColors == false ? options.dataLabelBkgrdColors : false
	if (['b', 'l', 'r', 't', 'tr'].indexOf(options.legendPos || '') < 0) options.legendPos = 'r'
	// barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
	if (['clustered', 'standard', 'stacked', 'percentStacked'].indexOf(options.barGrouping || '') < 0) options.barGrouping = 'standard'
	if (options.barGrouping.indexOf('tacked') > -1) {
		options.dataLabelPosition = 'ctr' // IMPORTANT: PPT-Online will not open Presentation when 'outEnd' etc is used on stacked!
		if (!options.barGapWidthPct) options.barGapWidthPct = 50
	}
	// 3D bar: ST_Shape
	if (['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax'].indexOf(options.bar3DShape || '') < 0) options.bar3DShape = 'box'
	// lineDataSymbol: http://www.datypic.com/sc/ooxml/a-val-32.html
	// Spec has [plus,star,x] however neither PPT2013 nor PPT-Online support them
	if (['circle', 'dash', 'diamond', 'dot', 'none', 'square', 'triangle'].indexOf(options.lineDataSymbol || '') < 0) options.lineDataSymbol = 'circle'
	if (['gap', 'span'].indexOf(options.displayBlanksAs || '') < 0) options.displayBlanksAs = 'span'
	if (['standard', 'marker', 'filled'].indexOf(options.radarStyle || '') < 0) options.radarStyle = 'standard'
	options.lineDataSymbolSize = options.lineDataSymbolSize && !isNaN(options.lineDataSymbolSize) ? options.lineDataSymbolSize : 6
	options.lineDataSymbolLineSize = options.lineDataSymbolLineSize && !isNaN(options.lineDataSymbolLineSize) ? options.lineDataSymbolLineSize * ONEPT : 0.75 * ONEPT
	// `layout` allows the override of PPT defaults to maximize space
	if (options.layout) {
		;['x', 'y', 'w', 'h'].forEach(key => {
			var val = options.layout[key]
			if (isNaN(Number(val)) || val < 0 || val > 1) {
				console.warn('Warning: chart.layout.' + key + ' can only be 0-1')
				delete options.layout[key] // remove invalid value so that default will be used
			}
		})
	}

	// Set gridline defaults
	options.catGridLine = options.catGridLine || (options.type == CHART_TYPES.SCATTER ? { color: 'D9D9D9', size: 1 } : {style:'none'})
	options.valGridLine = options.valGridLine || (options.type == CHART_TYPES.SCATTER ? { color: 'D9D9D9', size: 1 } : {})
	options.serGridLine = options.serGridLine || (options.type == CHART_TYPES.SCATTER ? { color: 'D9D9D9', size: 1 } : {style:'none'})
	correctGridLineOptions(options.catGridLine)
	correctGridLineOptions(options.valGridLine)
	correctGridLineOptions(options.serGridLine)
	correctShadowOptions(options.shadow)

	// C: Options: plotArea
	options.showDataTable = options.showDataTable == true || options.showDataTable == false ? options.showDataTable : false
	options.showDataTableHorzBorder = options.showDataTableHorzBorder == true || options.showDataTableHorzBorder == false ? options.showDataTableHorzBorder : true
	options.showDataTableVertBorder = options.showDataTableVertBorder == true || options.showDataTableVertBorder == false ? options.showDataTableVertBorder : true
	options.showDataTableOutline = options.showDataTableOutline == true || options.showDataTableOutline == false ? options.showDataTableOutline : true
	options.showDataTableKeys = options.showDataTableKeys == true || options.showDataTableKeys == false ? options.showDataTableKeys : true
	options.showLabel = options.showLabel == true || options.showLabel == false ? options.showLabel : false
	options.showLegend = options.showLegend == true || options.showLegend == false ? options.showLegend : false
	options.showPercent = options.showPercent == true || options.showPercent == false ? options.showPercent : true
	options.showTitle = options.showTitle == true || options.showTitle == false ? options.showTitle : false
	options.showValue = options.showValue == true || options.showValue == false ? options.showValue : false
	options.catAxisLineShow = typeof options.catAxisLineShow !== 'undefined' ? options.catAxisLineShow : true
	options.valAxisLineShow = typeof options.valAxisLineShow !== 'undefined' ? options.valAxisLineShow : true
	options.serAxisLineShow = typeof options.serAxisLineShow !== 'undefined' ? options.serAxisLineShow : true

	options.v3DRotX = !isNaN(options.v3DRotX) && options.v3DRotX >= -90 && options.v3DRotX <= 90 ? options.v3DRotX : 30
	options.v3DRotY = !isNaN(options.v3DRotY) && options.v3DRotY >= 0 && options.v3DRotY <= 360 ? options.v3DRotY : 30
	options.v3DRAngAx = options.v3DRAngAx == true || options.v3DRAngAx == false ? options.v3DRAngAx : true
	options.v3DPerspective = !isNaN(options.v3DPerspective) && options.v3DPerspective >= 0 && options.v3DPerspective <= 240 ? options.v3DPerspective : 30

	// D: Options: chart
	options.barGapWidthPct = !isNaN(options.barGapWidthPct) && options.barGapWidthPct >= 0 && options.barGapWidthPct <= 1000 ? options.barGapWidthPct : 150
	options.barGapDepthPct = !isNaN(options.barGapDepthPct) && options.barGapDepthPct >= 0 && options.barGapDepthPct <= 1000 ? options.barGapDepthPct : 150

	options.chartColors = Array.isArray(options.chartColors)
		? options.chartColors
		: options.type == CHART_TYPES.PIE || options.type == CHART_TYPES.DOUGHNUT
		? PIECHART_COLORS
		: BARCHART_COLORS
	options.chartColorsOpacity = options.chartColorsOpacity && !isNaN(options.chartColorsOpacity) ? options.chartColorsOpacity : null
	//
	options.border = options.border && typeof options.border === 'object' ? options.border : null
	if (options.border && (!options.border.pt || isNaN(options.border.pt))) options.border.pt = 1
	if (options.border && (!options.border.color || typeof options.border.color !== 'string' || options.border.color.length != 6)) options.border.color = '363636'
	//
	options.dataBorder = options.dataBorder && typeof options.dataBorder === 'object' ? options.dataBorder : null
	if (options.dataBorder && (!options.dataBorder.pt || isNaN(options.dataBorder.pt))) options.dataBorder.pt = 0.75
	if (options.dataBorder && (!options.dataBorder.color || typeof options.dataBorder.color !== 'string' || options.dataBorder.color.length != 6))
		options.dataBorder.color = 'F9F9F9'
	//
	if (!options.dataLabelFormatCode && options.type === CHART_TYPES.SCATTER) options.dataLabelFormatCode = 'General'
	options.dataLabelFormatCode =
		options.dataLabelFormatCode && typeof options.dataLabelFormatCode === 'string'
			? options.dataLabelFormatCode
			: options.type == CHART_TYPES.PIE || options.type == CHART_TYPES.DOUGHNUT
			? '0%'
			: '#,##0'
	//
	// Set default format for Scatter chart labels to custom string if not defined
	if (!options.dataLabelFormatScatter && options.type === CHART_TYPES.SCATTER) options.dataLabelFormatScatter = 'custom'
	//
	options.lineSize = typeof options.lineSize === 'number' ? options.lineSize : 2
	options.valAxisMajorUnit = typeof options.valAxisMajorUnit === 'number' ? options.valAxisMajorUnit : null
	options.valAxisCrossesAt = options.valAxisCrossesAt || 'autoZero'

	// STEP 4: Set props
	resultObject.type = 'chart'
	resultObject.options = options

	// STEP 5: Add this chart to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	targetRels.push({
		rId: chartRelId,
		data: tmpData,
		opts: options,
		type: SLIDE_OBJECT_TYPES.chart,
		globalId: chartId,
		fileName: 'chart' + chartId + '.xml',
		Target: '/ppt/charts/chart' + chartId + '.xml',
	})
	resultObject.chartRid = chartRelId

	target.data.push(resultObject)
	return resultObject
}

/* ===== */

/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * The following object is expected as a slide definition:
 * {
 *   bkgd: 'FF00FF',
 *   objects: [{
 *     text: {
 *       text: 'Hello World',
 *       x: 1,
 *       y: 1
 *     }
 *   }]
 * }
 * @param {Object} `slideDef` slide definition
 * @param {ISlide} `target` empty slide object that should be updated by the passed definition
 */
export function createSlideObject(slideDef, target) {
	// STEP 1: Add background
	if (slideDef.bkgd) {
		addBackgroundDefinition(slideDef.bkgd, target)
	}

	// STEP 2: Add all Slide Master objects in the order they were given (Issue#53)
	if (slideDef.objects && Array.isArray(slideDef.objects) && slideDef.objects.length > 0) {
		slideDef.objects.forEach((object, idx:number) => {
			var key = Object.keys(object)[0]
			if (MASTER_OBJECTS[key] && key == 'chart') addChartDefinition(object.chart.type, object.chart.data, object.chart.opts, target)
			else if (MASTER_OBJECTS[key] && key == 'image') addImageDefinition(object[key], target)
			else if (MASTER_OBJECTS[key] && key == 'line') addShapeDefinition(BASE_SHAPES.LINE, object[key], target)
			else if (MASTER_OBJECTS[key] && key == 'rect') addShapeDefinition(BASE_SHAPES.RECTANGLE, object[key], target)
			else if (MASTER_OBJECTS[key] && key == 'text') addTextDefinition(object[key].text, object[key].options, target, false)
			else if (MASTER_OBJECTS[key] && key == 'placeholder') {
				// TODO: 20180820: Check for existing `name`?
				object[key].options.placeholderName = object[key].options.name
				delete object[key].options.name // remap name for earier handling internally
				object[key].options.placeholderType = object[key].options.type
				delete object[key].options.type // remap name for earier handling internally
				object[key].options.placeholderIdx = 100 + idx
				addPlaceholderDefinition(object[key].text, object[key].options, target)
			}
		})
	}

	// STEP 3: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
	if (slideDef.slideNumber && typeof slideDef.slideNumber === 'object') {
		target.slideNumberObj = slideDef.slideNumber
	}
}

/**
 * Transforms a slide or slideLayout to resulting XML string.
 * @param {ISlide|ISlideLayout} slideObject slide object created within createSlideObject
 * @return {string} XML string with <p:cSld> as the root
 */
export function slideObjectToXml(slideObject: ISlide | ISlideLayout): string {
	let strSlideXml: string = slideObject.name ? '<p:cSld name="' + slideObject.name + '">' : '<p:cSld>'
	let intTableNum: number = 1

	// STEP 1: Add background
	if (slideObject && slideObject['bkgd']) {
		strSlideXml += genXmlColorSelection(false, slideObject['bkgd'])
	}

	// STEP 2: Add background image (using Strech) (if any)
	if (slideObject && slideObject['bkgdImgRid']) {
		// FIXME: We should be doing this in the slideLayout...
		strSlideXml +=
			'<p:bg>' +
			'<p:bgPr><a:blipFill dpi="0" rotWithShape="1">' +
			'<a:blip r:embed="rId' +
			slideObject['bkgdImgRid'] +
			'"><a:lum/></a:blip>' +
			'<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>' +
			'<a:effectLst/></p:bgPr>' +
			'</p:bg>'
	}

	// STEP 3: Continue slide by starting spTree node
	strSlideXml += '<p:spTree>'
	strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
	strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
	strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'

	// STEP 4: Loop over all Slide.data objects and add them to this slide ===============================
	slideObject.data.forEach((slideItemObj, idx: number) => {
		let x = 0,
			y = 0,
			cx = getSmartParseNumber('75%', 'X', slideObject['slideLayout'] || slideObject),
			cy = 0
		let placeholderObj: ISlideLayoutData
		let locationAttr = '',
			shapeType = null

		if (slideObject['slideLayout'] && slideObject['slideLayout']['data'] && slideItemObj.options && slideItemObj.options.placeholder) {
			placeholderObj = slideObject['slideLayout']['data'].filter((slideLayout: ISlideLayoutData) => {
				return slideLayout.options.placeholderName == slideItemObj.options.placeholder
			})[0]
		}

		// A: Set option vars
		slideItemObj.options = slideItemObj.options || {}

		if (slideItemObj.options.w || slideItemObj.options.w == 0) slideItemObj.options.cx = slideItemObj.options.w
		if (slideItemObj.options.h || slideItemObj.options.h == 0) slideItemObj.options.cy = slideItemObj.options.h
		//
		if (slideItemObj.options.x || slideItemObj.options.x == 0) x = getSmartParseNumber(slideItemObj.options.x, 'X', slideObject['slideLayout'] || slideObject)
		if (slideItemObj.options.y || slideItemObj.options.y == 0) y = getSmartParseNumber(slideItemObj.options.y, 'Y', slideObject['slideLayout'] || slideObject)
		if (slideItemObj.options.cx || slideItemObj.options.cx == 0) cx = getSmartParseNumber(slideItemObj.options.cx, 'X', slideObject['slideLayout'] || slideObject)
		if (slideItemObj.options.cy || slideItemObj.options.cy == 0) cy = getSmartParseNumber(slideItemObj.options.cy, 'Y', slideObject['slideLayout'] || slideObject)

		// If using a placeholder then inherit it's position
		if (placeholderObj) {
			if (placeholderObj.options.x || placeholderObj.options.x == 0) x = getSmartParseNumber(placeholderObj.options.x, 'X', slideObject['slideLayout'] || slideObject)
			if (placeholderObj.options.y || placeholderObj.options.y == 0) y = getSmartParseNumber(placeholderObj.options.y, 'Y', slideObject['slideLayout'] || slideObject)
			if (placeholderObj.options.cx || placeholderObj.options.cx == 0)
				cx = getSmartParseNumber(placeholderObj.options.cx, 'X', slideObject['slideLayout'] || slideObject)
			if (placeholderObj.options.cy || placeholderObj.options.cy == 0)
				cy = getSmartParseNumber(placeholderObj.options.cy, 'Y', slideObject['slideLayout'] || slideObject)
		}
		//
		if (slideItemObj.options.shape) shapeType = getShapeInfo(slideItemObj.options.shape)
		//
		if (slideItemObj.options.flipH) locationAttr += ' flipH="1"'
		if (slideItemObj.options.flipV) locationAttr += ' flipV="1"'
		if (slideItemObj.options.rotate) locationAttr += ' rot="' + convertRotationDegrees(slideItemObj.options.rotate) + '"'

		// B: Add OBJECT to current Slide ----------------------------
		switch (slideItemObj.type) {
			case SLIDE_OBJECT_TYPES.table:
				// FIRST: Ensure we have rows - otherwise, bail!
				if (!slideItemObj.arrTabRows || (Array.isArray(slideItemObj.arrTabRows) && slideItemObj.arrTabRows.length == 0)) break

				// Set table vars
				var objTableGrid = {}
				var arrTabRows = slideItemObj.arrTabRows
				var objTabOpts = slideItemObj.options
				var intColCnt = 0,
					intColW = 0

				// Calc number of columns
				// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
				// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
				arrTabRows[0].forEach(cell => {
					var cellOpts = cell.options || null
					intColCnt += cellOpts && cellOpts.colspan ? Number(cellOpts.colspan) : 1
				})

				// STEP 1: Start Table XML =============================
				// NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
				let strXml =
					'<p:graphicFrame>' +
					'  <p:nvGraphicFramePr>' +
					'    <p:cNvPr id="' +
					(intTableNum * slideObject['number'] + 1) +
					'" name="Table ' +
					intTableNum * slideObject['number'] +
					'"/>' +
					'    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
					'    <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>' +
					'  </p:nvGraphicFramePr>' +
					'  <p:xfrm>' +
					'    <a:off  x="' +
					(x || (x == 0 ? 0 : EMU)) +
					'"  y="' +
					(y || (y == 0 ? 0 : EMU)) +
					'"/>' +
					'    <a:ext cx="' +
					(cx || (cx == 0 ? 0 : EMU)) +
					'" cy="' +
					(cy || (cy == 0 ? 0 : EMU)) +
					'"/>' +
					'  </p:xfrm>' +
					'  <a:graphic>' +
					'    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
					'      <a:tbl>' +
					'        <a:tblPr/>'
				// + '        <a:tblPr bandRow="1"/>';

				// FIXME: Support banded rows, first/last row, etc.
				// NOTE: Banding, etc. only shows when using a table style! (or set alt row color if banding)
				// <a:tblPr firstCol="0" firstRow="0" lastCol="0" lastRow="0" bandCol="0" bandRow="1">

				// STEP 2: Set column widths
				// Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
				// A: Col widths provided?
				if (Array.isArray(objTabOpts.colW)) {
					strXml += '<a:tblGrid>'
					for (var col = 0; col < intColCnt; col++) {
						strXml += '  <a:gridCol w="' + Math.round(inch2Emu(objTabOpts.colW[col]) || slideItemObj.options.cx / intColCnt) + '"/>'
					}
					strXml += '</a:tblGrid>'
				}
				// B: Table Width provided without colW? Then distribute cols
				else {
					intColW = objTabOpts.colW ? objTabOpts.colW : EMU
					if (slideItemObj.options.cx && !objTabOpts.colW) intColW = Math.round(slideItemObj.options.cx / intColCnt) // FIX: Issue#12
					strXml += '<a:tblGrid>'
					for (var col = 0; col < intColCnt; col++) {
						strXml += '<a:gridCol w="' + intColW + '"/>'
					}
					strXml += '</a:tblGrid>'
				}

				// STEP 3: Build our row arrays into an actual grid to match the XML we will be building next (ISSUE #36)
				// Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
				// so a simple loop below in XML building wont suffice to build table correctly.
				// We have to build an actual grid now
				/*
						EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

						/------|------|------|------\
						|  A0  |  B0  |  C0  |  D0  |
						|      |  B1  |  C1  |      |
						|      |      |  C2  |  D2  |
						\------|------|------|------/
					*/
				arrTabRows.forEach((row, rIdx) => {
					// A: Create row if needed (recall one may be created in loop below for rowspans, so dont assume we need to create one each iteration)
					if (!objTableGrid[rIdx]) objTableGrid[rIdx] = {}

					// B: Loop over all cells
					row.forEach((cell, cIdx) => {
						// DESIGN: NOTE: Row cell arrays can be "uneven" (diff cell count in each) due to rowspan/colspan
						// Therefore, for each cell we run 0->colCount to determien the correct slot for it to reside
						// as the uneven/mixed nature of the data means we cannot use the cIdx value alone.
						// E.g.: the 2nd element in the row array may actually go into the 5th table grid row cell b/c of colspans!
						for (var idx = 0; cIdx + idx < intColCnt; idx++) {
							var currColIdx = cIdx + idx

							if (!objTableGrid[rIdx][currColIdx]) {
								// A: Set this cell
								objTableGrid[rIdx][currColIdx] = cell

								// B: Handle `colspan` or `rowspan` (a {cell} cant have both! FIXME: FUTURE: ROWSPAN & COLSPAN in same cell)
								if (cell && cell.opts && cell.opts.colspan && !isNaN(Number(cell.opts.colspan))) {
									for (var idy = 1; idy < Number(cell.opts.colspan); idy++) {
										objTableGrid[rIdx][currColIdx + idy] = { hmerge: true, text: 'hmerge' }
									}
								} else if (cell && cell.opts && cell.opts.rowspan && !isNaN(Number(cell.opts.rowspan))) {
									for (var idz = 1; idz < Number(cell.opts.rowspan); idz++) {
										if (!objTableGrid[rIdx + idz]) objTableGrid[rIdx + idz] = {}
										objTableGrid[rIdx + idz][currColIdx] = { vmerge: true, text: 'vmerge' }
									}
								}

								// C: Break out of colCnt loop now that slot has been filled
								break
							}
						}
					})
				})

				/* Only useful for rowspan/colspan testing
					if ( objTabOpts.debug ) {
						console.table(objTableGrid);
						var arrText = [];
						jQuery.each(objTableGrid, function(i,row){ var arrRow = []; jQuery.each(row,function(i,cell){ arrRow.push(cell.text); }); arrText.push(arrRow); });
						console.table( arrText );
					}
					*/

				// STEP 4: Build table rows/cells ============================
				jQuery.each(objTableGrid, (rIdx, rowObj) => {
					// A: Table Height provided without rowH? Then distribute rows
					var intRowH = 0 // IMPORTANT: Default must be zero for auto-sizing to work
					if (Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx]) intRowH = inch2Emu(Number(objTabOpts.rowH[rIdx]))
					else if (objTabOpts.rowH && !isNaN(Number(objTabOpts.rowH))) intRowH = inch2Emu(Number(objTabOpts.rowH))
					else if (slideItemObj.options.cy || slideItemObj.options.h)
						intRowH = (slideItemObj.options.h ? inch2Emu(slideItemObj.options.h) : slideItemObj.options.cy) / arrTabRows.length

					// B: Start row
					strXml += '<a:tr h="' + intRowH + '">'

					// C: Loop over each CELL
					jQuery.each(rowObj, (_cIdx, cell: ITableCell) => {
						// 1: "hmerge" cells are just place-holders in the table grid - skip those and go to next cell
						if (cell.hmerge) return

						// 2: OPTIONS: Build/set cell options ===========================
						{
							var cellOpts = cell.options || ({} as ITableCell['options'])
							/// TODO-3: FIXME: ONLY MAKE CELLS with objects! if (typeof cell === 'number' || typeof cell === 'string') cell = { text: cell.toString() }
							cellOpts.isTableCell = true // Used to create textBody XML
							cell.options = cellOpts

							// B: Apply default values (tabOpts being used when cellOpts dont exist):
							// SEE: http://officeopenxml.com/drwTableCellProperties-alignment.php
							;['align', 'bold', 'border', 'color', 'fill', 'fontFace', 'fontSize', 'margin', 'underline', 'valign'].forEach(name => {
								if (objTabOpts[name] && !cellOpts[name] && cellOpts[name] != 0) cellOpts[name] = objTabOpts[name]
							})

							var cellValign = cellOpts.valign
								? ' anchor="' +
								  cellOpts.valign
										.replace(/^c$/i, 'ctr')
										.replace(/^m$/i, 'ctr')
										.replace('center', 'ctr')
										.replace('middle', 'ctr')
										.replace('top', 't')
										.replace('btm', 'b')
										.replace('bottom', 'b') +
								  '"'
								: ''
							var cellColspan = cellOpts.colspan ? ' gridSpan="' + cellOpts.colspan + '"' : ''
							var cellRowspan = cellOpts.rowspan ? ' rowSpan="' + cellOpts.rowspan + '"' : ''
							var cellFill =
								(cell.optImp && cell.optImp.fill) || cellOpts.fill
									? ' <a:solidFill><a:srgbClr val="' + ((cell.optImp && cell.optImp.fill) || cellOpts.fill.replace('#', '')) + '"/></a:solidFill>'
									: ''
							var cellMargin = cellOpts.margin == 0 || cellOpts.margin ? cellOpts.margin : DEF_CELL_MARGIN_PT
							if (!Array.isArray(cellMargin) && typeof cellMargin === 'number') cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin]
							cellMargin =
								' marL="' +
								cellMargin[3] * ONEPT +
								'" marR="' +
								cellMargin[1] * ONEPT +
								'" marT="' +
								cellMargin[0] * ONEPT +
								'" marB="' +
								cellMargin[2] * ONEPT +
								'"'
						}

						// FIXME: Cell NOWRAP property (text wrap: add to a:tcPr (horzOverflow="overflow" or whatev opts exist)

						// 3: ROWSPAN: Add dummy cells for any active rowspan
						if (cell.vmerge) {
							strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>'
							return
						}

						// 4: Set CELL content and properties ==================================
						// FIXME: cell is "0" ??? table demo
						console.log(rowObj)
						console.log(cell)
						strXml += '<a:tc' + cellColspan + cellRowspan + '>' + genXmlTextBody(cell) + '<a:tcPr' + cellMargin + cellValign + '>'

						// 5: Borders: Add any borders
						/// TODO=3: FIXME: stop using `none` if (cellOpts.border && typeof cellOpts.border === 'string' && cellOpts.border.toLowerCase() == 'none') {
						if (cellOpts.border && cellOpts.border.type == 'none') {
							strXml += '  <a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>'
							strXml += '  <a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR>'
							strXml += '  <a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT>'
							strXml += '  <a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>'
						} else if (cellOpts.border && typeof cellOpts.border === 'string') {
							strXml +=
								'  <a:lnL w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnL>'
							strXml +=
								'  <a:lnR w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnR>'
							strXml +=
								'  <a:lnT w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnT>'
							strXml +=
								'  <a:lnB w="' + ONEPT + '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' + cellOpts.border + '"/></a:solidFill></a:lnB>'
						} else if (cellOpts.border && Array.isArray(cellOpts.border)) {
							jQuery.each([{ idx: 3, name: 'lnL' }, { idx: 1, name: 'lnR' }, { idx: 0, name: 'lnT' }, { idx: 2, name: 'lnB' }], (i, obj) => {
								if (cellOpts.border[obj.idx]) {
									var strC =
										'<a:solidFill><a:srgbClr val="' +
										(cellOpts.border[obj.idx].color ? cellOpts.border[obj.idx].color : DEF_CELL_BORDER.color) +
										'"/></a:solidFill>'
									var intW =
										cellOpts.border[obj.idx] && (cellOpts.border[obj.idx].pt || cellOpts.border[obj.idx].pt == 0)
											? ONEPT * Number(cellOpts.border[obj.idx].pt)
											: ONEPT
									strXml += '<a:' + obj.name + ' w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strC + '</a:' + obj.name + '>'
								} else strXml += '<a:' + obj.name + ' w="0"><a:miter lim="400000" /></a:' + obj.name + '>'
							})
						} else if (cellOpts.border && typeof cellOpts.border === 'object') {
							var intW = cellOpts.border && (cellOpts.border.pt || cellOpts.border.pt == 0) ? ONEPT * Number(cellOpts.border.pt) : ONEPT
							var strClr =
								'<a:solidFill><a:srgbClr val="' +
								(cellOpts.border.color ? cellOpts.border.color.replace('#', '') : DEF_CELL_BORDER.color) +
								'"/></a:solidFill>'
							var strAttr = '<a:prstDash val="'
							strAttr += cellOpts.border.type && cellOpts.border.type.toLowerCase().indexOf('dash') > -1 ? 'sysDash' : 'solid'
							strAttr += '"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>'
							// *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
							strXml += '<a:lnL w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnL>'
							strXml += '<a:lnR w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnR>'
							strXml += '<a:lnT w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnT>'
							strXml += '<a:lnB w="' + intW + '" cap="flat" cmpd="sng" algn="ctr">' + strClr + strAttr + '</a:lnB>'
							// *** IMPORTANT! *** LRTB order matters!
						}

						// 6: Close cell Properties & Cell
						strXml += cellFill
						strXml += '  </a:tcPr>'
						strXml += ' </a:tc>'

						// LAST: COLSPAN: Add a 'merged' col for each column being merged (SEE: http://officeopenxml.com/drwTableGrid.php)
						if (cellOpts.colspan) {
							for (var tmp = 1; tmp < Number(cellOpts.colspan); tmp++) {
								strXml += '<a:tc hMerge="1"><a:tcPr/></a:tc>'
							}
						}
					})

					// D: Complete row
					strXml += '</a:tr>'
				})

				// STEP 5: Complete table
				strXml += '      </a:tbl>'
				strXml += '    </a:graphicData>'
				strXml += '  </a:graphic>'
				strXml += '</p:graphicFrame>'

				// STEP 6: Set table XML
				strSlideXml += strXml

				// LAST: Increment counter
				intTableNum++
				break

			case SLIDE_OBJECT_TYPES.text:
			case SLIDE_OBJECT_TYPES.placeholder:
				// Lines can have zero cy, but text should not
				if (!slideItemObj.options.line && cy == 0) cy = EMU * 0.3

				// Margin/Padding/Inset for textboxes
				if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
					slideItemObj.options.bodyProp.lIns = slideItemObj.options.margin[0] * ONEPT || 0
					slideItemObj.options.bodyProp.rIns = slideItemObj.options.margin[1] * ONEPT || 0
					slideItemObj.options.bodyProp.bIns = slideItemObj.options.margin[2] * ONEPT || 0
					slideItemObj.options.bodyProp.tIns = slideItemObj.options.margin[3] * ONEPT || 0
				} else if ((slideItemObj.options.margin || slideItemObj.options.margin == 0) && !isNaN(slideItemObj.options.margin)) {
					slideItemObj.options.bodyProp.lIns = slideItemObj.options.margin * ONEPT
					slideItemObj.options.bodyProp.rIns = slideItemObj.options.margin * ONEPT
					slideItemObj.options.bodyProp.bIns = slideItemObj.options.margin * ONEPT
					slideItemObj.options.bodyProp.tIns = slideItemObj.options.margin * ONEPT
				}

				if (shapeType == null) shapeType = getShapeInfo(null)

				// A: Start SHAPE =======================================================
				strSlideXml += '<p:sp>'

				// B: The addition of the "txBox" attribute is the sole determiner of if an object is a Shape or Textbox
				strSlideXml += '<p:nvSpPr><p:cNvPr id="' + (idx + 2) + '" name="Object ' + (idx + 1) + '"/>'
				strSlideXml += '<p:cNvSpPr' + (slideItemObj.options && slideItemObj.options.isTextBox ? ' txBox="1"/>' : '/>')
				strSlideXml += '<p:nvPr>'
				strSlideXml += slideItemObj.type === 'placeholder' ? genXmlPlaceholder(slideItemObj) : genXmlPlaceholder(placeholderObj)
				strSlideXml += '</p:nvPr>'
				strSlideXml += '</p:nvSpPr><p:spPr>'
				strSlideXml += '<a:xfrm' + locationAttr + '>'
				strSlideXml += '<a:off x="' + x + '" y="' + y + '"/>'
				strSlideXml += '<a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>'
				strSlideXml +=
					'<a:prstGeom prst="' +
					shapeType.name +
					'"><a:avLst>' +
					(slideItemObj.options.rectRadius
						? '<a:gd name="adj" fmla="val ' + Math.round((slideItemObj.options.rectRadius * EMU * 100000) / Math.min(cx, cy)) + '" />'
						: '') +
					'</a:avLst></a:prstGeom>'

				// Option: FILL
				strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>'

				// Shape Type: LINE: line color
				if (slideItemObj.options.line) {
					strSlideXml += '<a:ln' + (slideItemObj.options.lineSize ? ' w="' + slideItemObj.options.lineSize * ONEPT + '"' : '') + '>'
					strSlideXml += genXmlColorSelection(slideItemObj.options.line)
					if (slideItemObj.options.lineDash) strSlideXml += '<a:prstDash val="' + slideItemObj.options.lineDash + '"/>'
					if (slideItemObj.options.lineHead) strSlideXml += '<a:headEnd type="' + slideItemObj.options.lineHead + '"/>'
					if (slideItemObj.options.lineTail) strSlideXml += '<a:tailEnd type="' + slideItemObj.options.lineTail + '"/>'
					strSlideXml += '</a:ln>'
				}

				// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
				if (slideItemObj.options.shadow) {
					slideItemObj.options.shadow.type = slideItemObj.options.shadow.type || 'outer'
					slideItemObj.options.shadow.blur = (slideItemObj.options.shadow.blur || 8) * ONEPT
					slideItemObj.options.shadow.offset = (slideItemObj.options.shadow.offset || 4) * ONEPT
					slideItemObj.options.shadow.angle = (slideItemObj.options.shadow.angle || 270) * 60000
					slideItemObj.options.shadow.color = slideItemObj.options.shadow.color || '000000'
					slideItemObj.options.shadow.opacity = (slideItemObj.options.shadow.opacity || 0.75) * 100000

					strSlideXml += '<a:effectLst>'
					strSlideXml += '<a:' + slideItemObj.options.shadow.type + 'Shdw sx="100000" sy="100000" kx="0" ky="0" '
					strSlideXml += ' algn="bl" rotWithShape="0" blurRad="' + slideItemObj.options.shadow.blur + '" '
					strSlideXml += ' dist="' + slideItemObj.options.shadow.offset + '" dir="' + slideItemObj.options.shadow.angle + '">'
					strSlideXml += '<a:srgbClr val="' + slideItemObj.options.shadow.color + '">'
					strSlideXml += '<a:alpha val="' + slideItemObj.options.shadow.opacity + '"/></a:srgbClr>'
					strSlideXml += '</a:outerShdw>'
					strSlideXml += '</a:effectLst>'
				}

				/* FIXME: FUTURE: Text wrapping (copied from MS-PPTX export)
					// Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
					if ( slideItemObj.options.textWrap ) {
						strSlideXml += '<a:extLst>'
									+ '<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">'
									+ '<ma14:wrappingTextBoxFlag xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main" val="1" />'
									+ '</a:ext>'
									+ '</a:extLst>';
					}
					*/

				// B: Close Shape Properties
				strSlideXml += '</p:spPr>'

				// Add formatted text
				strSlideXml += genXmlTextBody(slideItemObj)

				// LAST: Close SHAPE =======================================================
				strSlideXml += '</p:sp>'
				break

			case SLIDE_OBJECT_TYPES.image:
				var sizing = slideItemObj.options.sizing,
					rounding = slideItemObj.options.rounding,
					width = cx,
					height = cy

				strSlideXml += '<p:pic>'
				strSlideXml += '  <p:nvPicPr>'
				strSlideXml += '    <p:cNvPr id="' + (idx + 2) + '" name="Object ' + (idx + 1) + '" descr="' + encodeXmlEntities(slideItemObj.image) + '">'
				if (slideItemObj.hyperlink && slideItemObj.hyperlink.url)
					strSlideXml +=
						'<a:hlinkClick r:id="rId' +
						slideItemObj.hyperlink.rId +
						'" tooltip="' +
						(slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') +
						'" />'
				if (slideItemObj.hyperlink && slideItemObj.hyperlink.slide)
					strSlideXml +=
						'<a:hlinkClick r:id="rId' +
						slideItemObj.hyperlink.rId +
						'" tooltip="' +
						(slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : '') +
						'" action="ppaction://hlinksldjump" />'
				strSlideXml += '    </p:cNvPr>'
				strSlideXml += '    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
				strSlideXml += '    <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>'
				strSlideXml += '  </p:nvPicPr>'
				strSlideXml += '<p:blipFill>'
				// NOTE: This works for both cases: either `path` or `data` contains the SVG
				if (
					(slideObject['relsMedia'] || []).filter(rel => {
						return rel.rId == slideItemObj.imageRid
					})[0] &&
					(slideObject['relsMedia'] || []).filter(rel => {
						return rel.rId == slideItemObj.imageRid
					})[0]['extn'] == 'svg'
				) {
					strSlideXml += '<a:blip r:embed="rId' + (slideItemObj.imageRid - 1) + '"/>'
					strSlideXml += '<a:extLst>'
					strSlideXml += '  <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
					strSlideXml += '    <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId' + slideItemObj.imageRid + '"/>'
					strSlideXml += '  </a:ext>'
					strSlideXml += '</a:extLst>'
				} else {
					strSlideXml += '<a:blip r:embed="rId' + slideItemObj.imageRid + '"/>'
				}
				if (sizing && sizing.type) {
					var boxW = sizing.w ? getSmartParseNumber(sizing.w, 'X', slideObject['slideLayout'] || slideObject) : cx,
						boxH = sizing.h ? getSmartParseNumber(sizing.h, 'Y', slideObject['slideLayout'] || slideObject) : cy,
						boxX = getSmartParseNumber(sizing.x || 0, 'X', slideObject['slideLayout'] || slideObject),
						boxY = getSmartParseNumber(sizing.y || 0, 'Y', slideObject['slideLayout'] || slideObject)

					strSlideXml += imageSizingXml[sizing.type]({ w: width, h: height }, { w: boxW, h: boxH, x: boxX, y: boxY })
					width = boxW
					height = boxH
				} else {
					strSlideXml += '  <a:stretch><a:fillRect/></a:stretch>'
				}
				strSlideXml += '</p:blipFill>'
				strSlideXml += '<p:spPr>'
				strSlideXml += ' <a:xfrm' + locationAttr + '>'
				strSlideXml += '  <a:off  x="' + x + '"  y="' + y + '"/>'
				strSlideXml += '  <a:ext cx="' + width + '" cy="' + height + '"/>'
				strSlideXml += ' </a:xfrm>'
				strSlideXml += ' <a:prstGeom prst="' + (rounding ? 'ellipse' : 'rect') + '"><a:avLst/></a:prstGeom>'
				strSlideXml += '</p:spPr>'
				strSlideXml += '</p:pic>'
				break

			case SLIDE_OBJECT_TYPES.media:
				if (slideItemObj.mtype == 'online') {
					strSlideXml += '<p:pic>'
					strSlideXml += ' <p:nvPicPr>'
					// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preview image rId, PowerPoint throws error!
					strSlideXml += ' <p:cNvPr id="' + (slideItemObj.mediaRid + 2) + '" name="Picture' + (idx + 1) + '"/>'
					strSlideXml += ' <p:cNvPicPr/>'
					strSlideXml += ' <p:nvPr>'
					strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>'
					strSlideXml += ' </p:nvPr>'
					strSlideXml += ' </p:nvPicPr>'
					// NOTE: `blip` is diferent than videos; also there's no preview "p:extLst" above but exists in videos
					strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' // NOTE: Preview image is required!
					strSlideXml += ' <p:spPr>'
					strSlideXml += '  <a:xfrm' + locationAttr + '>'
					strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>'
					strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>'
					strSlideXml += '  </a:xfrm>'
					strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
					strSlideXml += ' </p:spPr>'
					strSlideXml += '</p:pic>'
				} else {
					strSlideXml += '<p:pic>'
					strSlideXml += ' <p:nvPicPr>'
					// IMPORTANT: <p:cNvPr id="" value is critical - if not the same number as preiew image rId, PowerPoint throws error!
					strSlideXml +=
						' <p:cNvPr id="' +
						(slideItemObj.mediaRid + 2) +
						'" name="' +
						slideItemObj.media
							.split('/')
							.pop()
							.split('.')
							.shift() +
						'"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>'
					strSlideXml += ' <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
					strSlideXml += ' <p:nvPr>'
					strSlideXml += '  <a:videoFile r:link="rId' + slideItemObj.mediaRid + '"/>'
					strSlideXml += '  <p:extLst>'
					strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
					strSlideXml += '    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId' + (slideItemObj.mediaRid + 1) + '"/>'
					strSlideXml += '   </p:ext>'
					strSlideXml += '  </p:extLst>'
					strSlideXml += ' </p:nvPr>'
					strSlideXml += ' </p:nvPicPr>'
					strSlideXml += ' <p:blipFill><a:blip r:embed="rId' + (slideItemObj.mediaRid + 2) + '"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' // NOTE: Preview image is required!
					strSlideXml += ' <p:spPr>'
					strSlideXml += '  <a:xfrm' + locationAttr + '>'
					strSlideXml += '   <a:off x="' + x + '" y="' + y + '"/>'
					strSlideXml += '   <a:ext cx="' + cx + '" cy="' + cy + '"/>'
					strSlideXml += '  </a:xfrm>'
					strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
					strSlideXml += ' </p:spPr>'
					strSlideXml += '</p:pic>'
				}
				break

			case SLIDE_OBJECT_TYPES.chart:
				strSlideXml += '<p:graphicFrame>'
				strSlideXml += ' <p:nvGraphicFramePr>'
				strSlideXml += '   <p:cNvPr id="' + (idx + 2) + '" name="Chart ' + (idx + 1) + '"/>'
				strSlideXml += '   <p:cNvGraphicFramePr/>'
				strSlideXml += '   <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>'
				strSlideXml += ' </p:nvGraphicFramePr>'
				strSlideXml += ' <p:xfrm>'
				strSlideXml += '  <a:off  x="' + x + '"  y="' + y + '"/>'
				strSlideXml += '  <a:ext cx="' + cx + '" cy="' + cy + '"/>'
				strSlideXml += ' </p:xfrm>'
				strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
				strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
				strSlideXml += '   <c:chart r:id="rId' + slideItemObj.chartRid + '" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>'
				strSlideXml += '  </a:graphicData>'
				strSlideXml += ' </a:graphic>'
				strSlideXml += '</p:graphicFrame>'
				break
		}
	})

	// STEP 5: Add slide numbers last (if any)
	if (slideObject.slideNumberObj) {
		///if (!slideObject.slideNumberObj) slideObject.slideNumberObj = { x: 0.3, y: '90%' }

		strSlideXml +=
			'<p:sp>' +
			'  <p:nvSpPr>' +
			'    <p:cNvPr id="25" name="Slide Number Placeholder 24"/>' +
			'    <p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr>' +
			'    <p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr>' +
			'  </p:nvSpPr>' +
			'  <p:spPr>' +
			'    <a:xfrm>' +
			'      <a:off x="' +
			getSmartParseNumber(slideObject.slideNumberObj.x, 'X', slideObject['slideLayout'] || slideObject) +
			'" y="' +
			getSmartParseNumber(slideObject.slideNumberObj.y, 'Y', slideObject['slideLayout'] || slideObject) +
			'"/>' +
			'      <a:ext cx="' +
			(slideObject.slideNumberObj.w ? getSmartParseNumber(slideObject.slideNumberObj.w, 'X', slideObject['slideLayout'] || slideObject) : 800000) +
			'" cy="' +
			(slideObject.slideNumberObj.h ? getSmartParseNumber(slideObject.slideNumberObj.h, 'Y', slideObject['slideLayout'] || slideObject) : 300000) +
			'"/>' +
			'    </a:xfrm>' +
			'    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
			'    <a:extLst><a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}"><ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext></a:extLst>' +
			'  </p:spPr>'
		// ISSUE #68: "Page number styling"
		strSlideXml += '<p:txBody>'
		strSlideXml += '  <a:bodyPr/>'
		strSlideXml += '  <a:lstStyle><a:lvl1pPr>'
		if (slideObject.slideNumberObj.fontFace || slideObject.slideNumberObj.fontSize || slideObject.slideNumberObj.color) {
			strSlideXml += '<a:defRPr sz="' + (slideObject.slideNumberObj.fontSize ? Math.round(slideObject.slideNumberObj.fontSize) : '12') + '00">'
			if (slideObject.slideNumberObj.color) strSlideXml += genXmlColorSelection(slideObject.slideNumberObj.color)
			if (slideObject.slideNumberObj.fontFace)
				strSlideXml +=
					'<a:latin typeface="' +
					slideObject.slideNumberObj.fontFace +
					'"/><a:ea typeface="' +
					slideObject.slideNumberObj.fontFace +
					'"/><a:cs typeface="' +
					slideObject.slideNumberObj.fontFace +
					'"/>'
			strSlideXml += '</a:defRPr>'
		}
		strSlideXml += '</a:lvl1pPr></a:lstStyle>'
		strSlideXml += '<a:p><a:fld id="' + SLDNUMFLDID + '" type="slidenum">' + '<a:rPr lang="en-US" smtClean="0"/><a:t></a:t></a:fld>' + '<a:endParaRPr lang="en-US"/></a:p>'
		strSlideXml += '</p:txBody></p:sp>'
	}

	// STEP 6: Close spTree and finalize slide XML
	strSlideXml += '</p:spTree>'
	strSlideXml += '</p:cSld>'

	// LAST: Return
	return strSlideXml
}

/**
 * Transforms slide relations to XML string.
 * Extra relations that are not dynamic can be passed using the 2nd arg (e.g. theme relation in master file).
 * These relations use rId series that starts with 1-increased maximum of rIds used for dynamic relations.
 *
 * @param {ISlide} slideObject slide object whose relations are being transformed
 * @param {Object[]} defaultRels array of default relations (such objects expected: { target: <filepath>, type: <schemepath> })
 * @return {string} complete XML string ready to be saved as a file
 */
export function slideObjectRelationsToXml(slideObject: ISlide | ISlideLayout, defaultRels): string {
	var lastRid = 0 // stores maximum rId used for dynamic relations
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
	// Add any rels for this Slide (image/audio/video/youtube/chart)
	slideObject.rels.forEach((rel, idx) => {
		lastRid = Math.max(lastRid, rel.rId)
		if (rel.type.toLowerCase().indexOf('image') > -1) {
			strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>'
		} else if (rel.type.toLowerCase().indexOf('chart') > -1) {
			strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>'
		} else if (rel.type.toLowerCase().indexOf('audio') > -1) {
			// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
			if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
				strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>'
			else
				strXml +=
					'<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"/>'
		} else if (rel.type.toLowerCase().indexOf('video') > -1) {
			// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
			if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
				strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>'
			else
				strXml +=
					'<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>'
		} else if (rel.type.toLowerCase().indexOf('online') > -1) {
			// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
			if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
				strXml += '<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.microsoft.com/office/2007/relationships/image"/>'
			else
				strXml +=
					'<Relationship Id="rId' +
					rel.rId +
					'" Target="' +
					rel.Target +
					'" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>'
		} else if (rel.type.toLowerCase().indexOf('hyperlink') > -1) {
			if (rel.data == 'slide') {
				strXml +=
					'<Relationship Id="rId' +
					rel.rId +
					'" Target="slide' +
					rel.Target +
					'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>'
			} else {
				strXml +=
					'<Relationship Id="rId' +
					rel.rId +
					'" Target="' +
					rel.Target +
					'" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"/>'
			}
		} else if (rel.type.toLowerCase().indexOf('notesSlide') > -1) {
			strXml +=
				'<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>'
		}
	})

	defaultRels.forEach((rel, idx) => {
		strXml += '<Relationship Id="rId' + (lastRid + idx + 1) + '" Target="' + rel.target + '" Type="' + rel.type + '"/>'
	})

	strXml += '</Relationships>'
	return strXml
}

let imageSizingXml = {
	cover: function(imgSize, boxDim) {
		var imgRatio = imgSize.h / imgSize.w,
			boxRatio = boxDim.h / boxDim.w,
			isBoxBased = boxRatio > imgRatio,
			width = isBoxBased ? boxDim.h / imgRatio : boxDim.w,
			height = isBoxBased ? boxDim.h : boxDim.w * imgRatio,
			hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)),
			vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))
		return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '" /><a:stretch/>'
	},
	contain: function(imgSize, boxDim) {
		var imgRatio = imgSize.h / imgSize.w,
			boxRatio = boxDim.h / boxDim.w,
			widthBased = boxRatio > imgRatio,
			width = widthBased ? boxDim.w : boxDim.h / imgRatio,
			height = widthBased ? boxDim.w * imgRatio : boxDim.h,
			hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)),
			vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))
		return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '" /><a:stretch/>'
	},
	crop: function(imageSize, boxDim) {
		var l = boxDim.x,
			r = imageSize.w - (boxDim.x + boxDim.w),
			t = boxDim.y,
			b = imageSize.h - (boxDim.y + boxDim.h),
			lPerc = Math.round(1e5 * (l / imageSize.w)),
			rPerc = Math.round(1e5 * (r / imageSize.w)),
			tPerc = Math.round(1e5 * (t / imageSize.h)),
			bPerc = Math.round(1e5 * (b / imageSize.h))
		return '<a:srcRect l="' + lPerc + '" r="' + rPerc + '" t="' + tPerc + '" b="' + bPerc + '" /><a:stretch/>'
	},
}
