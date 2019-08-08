/**
 * PptxGenJS: Slide object generators
 */

import {
	EMU,
	ONEPT,
	MASTER_OBJECTS,
	BARCHART_COLORS,
	PIECHART_COLORS,
	DEF_CELL_MARGIN_PT,
	CHART_TYPES,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_SLIDE_MARGIN_IN,
	IMG_PLAYBTN,
	BASE_SHAPES,
	CHART_TYPE_NAMES,
	SLIDE_OBJECT_TYPES,
	TEXT_HALIGN,
	TEXT_VALIGN,
} from './core-enums'
import {
	ISlide,
	ITextOpts,
	ILayout,
	ISlideLayout,
	ISlideObject,
	IMediaOpts,
	IChartOpts,
	IChartMulti,
	IImageOpts,
	ITableCell,
	IText,
	Shape,
	ShapeOptions,
	TableOptions,
} from './core-interfaces'
import { getSmartParseNumber, inch2Emu } from './gen-utils'
import { correctShadowOptions, createHyperlinkRels, getSlidesForTableRows } from './gen-xml'

/** counter for included charts (used for index in their filenames) */
var _chartCounter: number = 0

/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {ISlideMasterDef} slideDef - slide definition
 * @param {ISlide|ISlideLayout} target - empty slide object that should be updated by the passed definition
 */
export function createSlideObject(slideDef /*:ISlideMasterDef*/, target /*FIXME :ISlide|ISlideLayout*/) {
	// STEP 1: Add background
	if (slideDef.bkgd) {
		addBackgroundDefinition(slideDef.bkgd, target)
	}

	// STEP 2: Add all Slide Master objects in the order they were given (Issue#53)
	if (slideDef.objects && Array.isArray(slideDef.objects) && slideDef.objects.length > 0) {
		slideDef.objects.forEach((object, idx: number) => {
			let key = Object.keys(object)[0]
			if (MASTER_OBJECTS[key] && key == 'chart') addChartDefinition(object.chart.type, object.chart.data, object.chart.opts, target)
			else if (MASTER_OBJECTS[key] && key == 'image') addImageDefinition(target, object[key])
			else if (MASTER_OBJECTS[key] && key == 'line') addShapeDefinition(target, BASE_SHAPES.LINE, object[key])
			else if (MASTER_OBJECTS[key] && key == 'rect') addShapeDefinition(target, BASE_SHAPES.RECTANGLE, object[key])
			else if (MASTER_OBJECTS[key] && key == 'text') addTextDefinition(target, object[key].text, object[key].options, false)
			else if (MASTER_OBJECTS[key] && key == 'placeholder') {
				// TODO: 20180820: Check for existing `name`?
				object[key].options.placeholder = object[key].options.name
				delete object[key].options.name // remap name for earier handling internally
				object[key].options.placeholderType = object[key].options.type
				delete object[key].options.type // remap name for earier handling internally
				object[key].options.placeholderIdx = 100 + idx
				addPlaceholderDefinition(target, object[key].text, object[key].options)
			}
		})
	}

	// STEP 3: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
	if (slideDef.slideNumber && typeof slideDef.slideNumber === 'object') {
		target.slideNumberObj = slideDef.slideNumber
	}
}

/**
 * Adds a background image or color to a slide definition.
 * @param {String|Object} bkg - color string or an object with image definition
 * @param {ISlide} target - slide object that the background is set to
 */
function addBackgroundDefinition(bkg: string | { src?: string; path?: string; data?: string }, target: ISlide) {
	if (typeof bkg === 'object' && (bkg.src || bkg.path || bkg.data)) {
		// Allow the use of only the data key (`path` isnt reqd)
		bkg.src = bkg.src || bkg.path || null
		if (!bkg.src) bkg.src = 'preencoded.png'
		let strImgExtn = (bkg.src.split('.').pop() || 'png').split('?')[0] // Handle "blah.jpg?width=540" etc.
		if (strImgExtn == 'jpg') strImgExtn = 'jpeg' // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup

		let intRels = target.relsMedia.length + 2 // `rId` needs to be >=2 as Id="rId1" is "SlideMaster1.xml"
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
 * Generate the chart based on input data.
 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
 *
 * @param {CHART_TYPE_NAMES | IChartMulti[]} `type` should belong to: 'column', 'pie'
 * @param {[]} `data` a JSON object with follow the following format
 * @param {IChartOpts} `opt` chart options
 * @param {ISlide} `target` slide object that the chart will be added to
 * @return {object} chart object
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
export function addChartDefinition(target: ISlide, type: CHART_TYPE_NAMES | IChartMulti[], data: [], opt: IChartOpts): object {
	function correctGridLineOptions(glOpts) {
		if (!glOpts || glOpts.style == 'none') return
		if (glOpts.size !== undefined && (isNaN(Number(glOpts.size)) || glOpts.size <= 0)) {
			console.warn('Warning: chart.gridLine.size must be greater than 0.')
			delete glOpts.size // delete prop to used defaults
		}
		if (glOpts.style && ['solid', 'dash', 'dot'].indexOf(glOpts.style) < 0) {
			console.warn('Warning: chart.gridLine.style options: `solid`, `dash`, `dot`.')
			delete glOpts.style
		}
	}

	let chartId = ++_chartCounter
	let resultObject = {
		type: null,
		text: null,
		options: null,
		chartRid: null,
	}
	// DESIGN: `type` can an object (ex: `pptx.charts.DOUGHNUT`) or an array of chart objects
	// EX: addChartDefinition([ { type:pptx.charts.BAR, data:{name:'', labels:[], values[]} }, {<etc>} ])
	// Multi-Type Charts
	let tmpOpt
	let tmpData = [],
		options: IChartOpts
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
	options.catGridLine = options.catGridLine || (options.type == CHART_TYPES.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' })
	options.valGridLine = options.valGridLine || (options.type == CHART_TYPES.SCATTER ? { color: 'D9D9D9', size: 1 } : {})
	options.serGridLine = options.serGridLine || (options.type == CHART_TYPES.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' })
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
	resultObject.chartRid = target.relsChart.length + 1

	// STEP 5: Add this chart to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	target.relsChart.push({
		rId: target.relsChart.length + 1,
		data: tmpData,
		opts: options,
		type: options.type,
		globalId: chartId,
		fileName: 'chart' + chartId + '.xml',
		Target: '/ppt/charts/chart' + chartId + '.xml',
	})

	target.data.push(resultObject)
	return resultObject
}

/**
 * Adds an image object to a slide definition.
 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
 * @param {IImageOpts} `opt` - object containing `path`/`data`, `x`, `y`, etc.
 * @param {ISlide} `target` - slide that the image should be added to (if not specified as the 2nd arg)
 * @return {Object} image object
 */
export function addImageDefinition(target: ISlide, opt: IImageOpts): object {
	let newObject: any = {
		type: null,
		text: null,
		options: null,
		image: null,
		imageRid: null,
		hyperlink: null,
	}
	// FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
	let intPosX = opt.x || 0
	let intPosY = opt.y || 0
	let intWidth = opt.w || 0
	let intHeight = opt.h || 0
	let sizing = opt.sizing || null
	let objHyperlink = opt.hyperlink || ''
	let strImageData = opt.data || ''
	let strImagePath = opt.path || ''
	let imageRelId = target.rels.length + target.relsChart.length + target.relsMedia.length + 1

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
	let strImgExtn =
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
	newObject.type = 'image'
	newObject.image = strImagePath || 'preencoded.png'

	// STEP 3: Set image properties & options
	// FIXME: Measure actual image when no intWidth/intHeight params passed
	// ....: This is an async process: we need to make getSizeFromImage use callback, then set H/W...
	// if ( !intWidth || !intHeight ) { var imgObj = getSizeFromImage(strImagePath);
	newObject.options = {
		x: intPosX || 0,
		y: intPosY || 0,
		w: intWidth || 1,
		h: intHeight || 1,
		rounding: typeof opt.rounding === 'boolean' ? opt.rounding : false,
		sizing: sizing,
		placeholder: opt.placeholder,
	}

	// STEP 4: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	if (strImgExtn == 'svg') {
		// SVG files consume *TWO* rId's: (a png version and the svg image)
		// <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
		// <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>
		target.relsMedia.push({
			path: strImagePath || strImageData + 'png',
			type: 'image/png',
			extn: 'png',
			data: strImageData || '',
			rId: imageRelId,
			Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.png',
			isSvgPng: true,
			svgSize: { w: newObject.options.w, h: newObject.options.h },
		})
		newObject.imageRid = imageRelId
		target.relsMedia.push({
			path: strImagePath || strImageData,
			type: 'image/' + strImgExtn,
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId + 1,
			Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.' + strImgExtn,
		})
		newObject.imageRid = imageRelId + 1
	} else {
		target.relsMedia.push({
			path: strImagePath || 'preencoded.' + strImgExtn,
			type: 'image/' + strImgExtn,
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId,
			Target: '../media/image-' + target.number + '-' + (target.relsMedia.length + 1) + '.' + strImgExtn,
		})
		newObject.imageRid = imageRelId
	}

	// STEP 5: Hyperlink support
	if (typeof objHyperlink === 'object') {
		if (!objHyperlink.url && !objHyperlink.slide) throw 'ERROR: `hyperlink` option requires either: `url` or `slide`'
		else {
			// TODO: 20190729: Why not use "createHyperlinkRels"?
			imageRelId++

			target.rels.push({
				type: SLIDE_OBJECT_TYPES.hyperlink,
				data: objHyperlink.slide ? 'slide' : 'dummy',
				rId: imageRelId,
				Target: objHyperlink.url || objHyperlink.slide.toString(),
			})

			objHyperlink.rId = imageRelId
			newObject.hyperlink = objHyperlink
		}
	}

	// STEP 6: Add object to slide
	target.data.push(newObject)

	// LAST
	return newObject
}

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
	if (strType == 'online') {
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
 * Adds a placeholder object to a slide definition.
 * @param {String} `text`
 * @param {Object} `opt`
 * @param {ISlide} `target` slide object that the placeholder should be added to
 */
export function addPlaceholderDefinition(target: ISlide, text: string, opt: object) {
	return addTextDefinition(target, text, opt, true)
}

/**
 * Adds a shape object to a slide definition.
 * @param {Shape} shape shape const object (pptx.shapes)
 * @param {ShapeOptions} opt
 * @param {ISlide} target slide object that the shape should be added to
 */
export function addShapeDefinition(target: ISlide, shape: Shape, opt: ShapeOptions) {
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
	options.x = options.x || (options.x == 0 ? 0 : 1)
	options.y = options.y || (options.y == 0 ? 0 : 1)
	options.w = options.w || (options.w == 0 ? 0 : 1)
	options.h = options.h || (options.h == 0 ? 0 : 1)
	options.line = options.line || (shape.name == 'line' ? '333333' : null)
	options.lineSize = options.lineSize || (shape.name == 'line' ? 1 : null)
	if (['dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'solid', 'sysDash', 'sysDot'].indexOf(options.lineDash || '') < 0) options.lineDash = 'solid'

	// 3: Add object to slide
	target.data.push(newObject)
}

/**
 * Adds a table object to a slide definition.
 * @param {ISlide} target - slide object that the table should be added to
 * @param {TODO} arrTabRows - table data
 * @param {TableOptions} inOpt - table options
 * @param {ISlideLayout} slideLayout - Slide layout
 * @param {ILayout} presLayout - Presenation layout
 */
export function addTableDefinition(target: ISlide, arrTabRows, inOpt: TableOptions, slideLayout: ISlideLayout, presLayout: ILayout, addSlide: Function, getSlide: Function) {
	let opt = inOpt && typeof inOpt === 'object' ? inOpt : ({} as TableOptions)
	let slides = [target] // Create array of Slides as more will be added for auto-paging

	// STEP 1: REALITY-CHECK
	if (arrTabRows == null || arrTabRows.length == 0 || !Array.isArray(arrTabRows)) {
		try {
			console.warn('addTable: Array expected! USAGE: slide.addTable( [rows], {options} );')
		} catch (ex) {}
		return null
	}

	// STEP 2: Row setup: Handle case where user passed in a simple 1-row array. EX: `["cell 1", "cell 2"]`
	//var arrRows = jQuery.extend(true,[],arrTabRows);
	//if ( !Array.isArray(arrRows[0]) ) arrRows = [ jQuery.extend(true,[],arrTabRows) ];
	let arrRows: [ITableCell[]] = arrTabRows as [ITableCell[]]
	if (!Array.isArray(arrRows[0])) arrRows = [arrTabRows]

	// STEP 3: Set options
	opt.x = getSmartParseNumber(opt.x || (opt.x == 0 ? 0 : EMU / 2), 'X', presLayout)
	opt.y = getSmartParseNumber(opt.y || (opt.y == 0 ? 0 : EMU), 'Y', presLayout)
	if (opt.h) opt.h = getSmartParseNumber(opt.h, 'Y', presLayout) // NOTE: Dont set default `h` - leaving it null triggers auto-rowH in `makeXMLSlide()`
	opt.autoPage = opt.autoPage == false ? false : true
	opt.fontSize = opt.fontSize || DEF_FONT_SIZE
	opt.lineWeight = typeof opt.lineWeight !== 'undefined' && !isNaN(Number(opt.lineWeight)) ? Number(opt.lineWeight) : 0
	opt.margin = opt.margin == 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_PT
	if (typeof opt.margin === 'number') opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)]
	if (opt.lineWeight > 1) opt.lineWeight = 1
	else if (opt.lineWeight < -1) opt.lineWeight = -1
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
		} else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length != arrRows[0].length) {
			console.warn('addTable: colW.length != data.length! Defaulting to evenly distributed col widths.')

			var numColWidth = Math.floor((presLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3]) / arrRows[0].length)
			opt.colW = []
			for (var idx = 0; idx < arrRows[0].length; idx++) {
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

	// STEP 5: Loop over cells: transform to ITableCell; check to see whether to skip autopaging
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
				else if (typeof cell.text === 'undefined' || cell.text == null) row[idy].text = ''

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

	// STEP 6: Create hyperlink rels
	createHyperlinkRels(this.slides, arrRows, target.rels)
	//console.log(this.slides)
	// FIXME: TODO-3: "this." refs above dont exist here!! 20190725
	// FIXME: why do we need all slides???

	// STEP 7: Auto-Paging: (via {options} and used internally)
	// (used internally by `tableToSlides()` to not engage recursion - we've already paged the table data, just add this one)
	if (opt && opt.autoPage == false) {
		// Add data (NOTE: Use `extend` to avoid mutation)
		target.data.push({
			type: SLIDE_OBJECT_TYPES.table,
			arrTabRows: arrRows,
			options: Object.assign({}, opt),
		})
	} else {
		// Loop over rows and create 1-N tables as needed (ISSUE#21)
		getSlidesForTableRows(arrRows, opt, presLayout, slideLayout).forEach((arrRows, idx) => {
			// A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
			if (!getSlide(target.number + idx)) slides.push(addSlide(presLayout ? presLayout.name : null))

			// B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
			if (idx > 0) opt.y = inch2Emu(opt.newSlideStartY || arrTableMargin[0])

			// C: Add this table to new Slide
			opt.autoPage = false
			getSlide(target.number + idx).addTable(arrRows, Object.assign({}, opt))
		})
	}
}

/**
 * Adds a text object to a slide definition.
 * @param {string|IText[]} text
 * @param {ITextOpts} opt
 * @param {ISlide} target - slide object that the text should be added to
 * @param {boolean} isPlaceholder` is this a placeholder object
 * @since: 1.0.0
 */
export function addTextDefinition(target: ISlide, text: string | IText[], opts: ITextOpts, isPlaceholder: boolean) {
	let opt: ITextOpts = opts || {}
	if (!opt.bodyProp) opt.bodyProp = {}
	let newObject = {
		text: (Array.isArray(text) && text.length == 0 ? '' : text || '') || '',
		type: isPlaceholder ? SLIDE_OBJECT_TYPES.placeholder : SLIDE_OBJECT_TYPES.text,
		options: opts,
		shape: opts.shape,
	}

	// STEP 2: Set some options
	{
		// A: Placeholders should inherit their colors or override them, so don't default them
		if (!opt.placeholder) {
			opt.color = opt.color || target.color || DEF_FONT_COLOR // Set color (options > inherit from Slide > default to black)
		}

		// B
		if (opt.shape && opt.shape.name == 'line') {
			opt.line = opt.line || '333333'
			opt.lineSize = opt.lineSize || 1
		}

		// C
		newObject.options.lineSpacing = opt.lineSpacing && !isNaN(opt.lineSpacing) ? opt.lineSpacing : null

		// D: Transform text options to bodyProperties as thats how we build XML
		newObject.options.bodyProp.autoFit = opt.autoFit || false // If true, shape will collapse to text size (Fit To Shape)
		newObject.options.bodyProp.anchor = !opt.placeholder ? TEXT_VALIGN.ctr : null // VALS: [t,ctr,b]
		newObject.options.bodyProp.vert = opt.vert || null // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]

		if ((opt.inset && !isNaN(Number(opt.inset))) || opt.inset == 0) {
			newObject.options.bodyProp.lIns = inch2Emu(opt.inset)
			newObject.options.bodyProp.rIns = inch2Emu(opt.inset)
			newObject.options.bodyProp.tIns = inch2Emu(opt.inset)
			newObject.options.bodyProp.bIns = inch2Emu(opt.inset)
		}
	}

	// STEP 3: Transform `align`/`valign` to XML values, store in bodyProp for XML gen
	{
		if ((newObject.options.align || '').toLowerCase().startsWith('c')) newObject.options.bodyProp.align = TEXT_HALIGN.center
		else if ((newObject.options.align || '').toLowerCase().startsWith('l')) newObject.options.bodyProp.align = TEXT_HALIGN.left
		else if ((newObject.options.align || '').toLowerCase().startsWith('r')) newObject.options.bodyProp.align = TEXT_HALIGN.right
		else if ((newObject.options.align || '').toLowerCase().startsWith('j')) newObject.options.bodyProp.align = TEXT_HALIGN.justify

		if ((newObject.options.valign || '').toLowerCase().startsWith('b')) newObject.options.bodyProp.anchor = TEXT_VALIGN.b
		else if ((newObject.options.valign || '').toLowerCase().startsWith('c')) newObject.options.bodyProp.anchor = TEXT_VALIGN.ctr
		else if ((newObject.options.valign || '').toLowerCase().startsWith('t')) newObject.options.bodyProp.anchor = TEXT_VALIGN.t
	}

	// STEP 4: ROBUST: Set rational values for some shadow props if needed
	correctShadowOptions(opt.shadow)

	// STEP 5: Create hyperlinks
	createHyperlinkRels([target], newObject.text || '', target.rels)

	// LAST: Add object to Slide
	target.data.push(newObject)
}

export function addPlaceholdersToSlideLayouts(slide: ISlide) {
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
				addTextDefinition(slide, '', { placeholder: slideLayoutObj.options.placeholder }, false)
			}
		}
	})
}
