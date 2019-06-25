/**
 * PptxGenJS: XML Generation
 */

import {
	CRLF,
	EMU,
	ONEPT,
	MASTER_OBJECTS,
	REGEX_HEX_COLOR,
	BARCHART_COLORS,
	PIECHART_COLORS,
	DEF_CELL_BORDER,
	DEF_CELL_MARGIN_PT,
	DEF_CHART_GRIDLINE,
	DEF_SHAPE_SHADOW,
	SCHEME_COLOR_NAMES,
	CHART_TYPES,
	LETTERS,
	SLDNUMFLDID,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY,
	AXIS_ID_SERIES_PRIMARY,
	BULLET_TYPES,
	DEF_FONT_TITLE_SIZE,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	LAYOUT_IDX_SERIES_BASE,
	PLACEHOLDER_TYPES,
	SLIDE_OBJECT_TYPES,
} from './enums'
import { IChartOpts, ISlide, ISlideRelChart, IShadowOpts, ITextOpts, ILayout, ISlideLayout, ISlideDataObject, ITableCell } from './interfaces'
import { convertRotationDegrees, getSmartParseNumber, encodeXmlEntities, getMix, getUuid, inch2Emu } from './utils'
import { gObjPptxShapes } from './shapes'
//import jszip from 'jszip'
import { each } from 'jquery'
import * as JSZip from 'jszip'

/** counter for included images (used for index in their filenames) */
var _imageCounter: number = 0
/** counter for included charts (used for index in their filenames) */
var _chartCounter: number = 0

export var gObjPptxGenerators = {
	/**
	 * Adds a background image or color to a slide definition.
	 * @param {String|Object} bkg color string or an object with image definition
	 * @param {ISlide} target slide object that the background is set to
	 */
	addBackgroundDefinition: function addBackgroundDefinition(bkg: string | { src?: string; path?: string; data?: string }, target: ISlide) {
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
			target.back = bkg
		}
	},

	/**
	 * Adds a text object to a slide definition.
	 * @param {String} text
	 * @param {ITextOpts} opt
	 * @param {ISlide} target - slide object that the text should be added to
	 * @param {Boolean} isPlaceholder
	 * @since: 1.0.0
	 */
	addTextDefinition: function addTextDefinition(text: string | Array<object>, opt: ITextOpts, target: ISlide, isPlaceholder: boolean) {
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
	},

	/**
	 * Adds Notes to a slide.
	 * @param {String} `notes`
	 * @param {Object} opt (*unused*)
	 * @param {ISlide} `target` slide object
	 * @since 2.3.0
	 */
	addNotesDefinition: function addNotesDefinition(notes: string, opt: object, target: ISlide) {
		var opt = opt && typeof opt === 'object' ? opt : {}
		var resultObject: ISlideDataObject = {
			type: null,
			text: null,
		}

		resultObject.type = SLIDE_OBJECT_TYPES.notes
		resultObject.text = notes

		target.data.push(resultObject)

		return resultObject
	},

	/**
	 * Adds a placeholder object to a slide definition.
	 * @param {String} `text`
	 * @param {Object} `opt`
	 * @param {ISlide} `target` slide object that the placeholder should be added to
	 */
	addPlaceholderDefinition: function addPlaceholderDefinition(text: string, opt: object, target: ISlide) {
		return gObjPptxGenerators.addTextDefinition(text, opt, target, true)
	},

	/**
	 * Adds a shape object to a slide definition.
	 * @param {gObjPptxShapes} shape shape const object (pptx.shapes)
	 * @param {Object} opt
	 * @param {Object} target slide object that the shape should be added to
	 * @return {Object} shape object
	 */
	addShapeDefinition: function addShapeDefinition(shape, opt, target) {
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
	},

	/**
	 * Adds an image object to a slide definition.
	 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
	 * @param {Object} objImage - object containing `path`/`data`, `x`, `y`, etc.
	 * @param {Object} target - slide that the image should be added to (if not specified as the 2nd arg)
	 * @return {Object} image object
	 */
	addImageDefinition: function addImageDefinition(objImage, target) {
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
	},

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
	addChartDefinition: function addChartDefinition(type, data, opt, target) {
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
			options
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
		options.x = typeof options.x !== 'undefined' && options.x != null && !isNaN(options.x) ? options.x : 1
		options.y = typeof options.y !== 'undefined' && options.y != null && !isNaN(options.y) ? options.y : 1
		options.w = options.w || '50%'
		options.h = options.h || '50%'

		// B: Options: misc
		if (['bar', 'col'].indexOf(options.barDir || '') < 0) options.barDir = 'col'
		// IMPORTANT: 'bestFit' will cause issues with PPT-Online in some cases, so defualt to 'ctr'!
		if (['bestFit', 'b', 'ctr', 'inBase', 'inEnd', 'l', 'outEnd', 'r', 't'].indexOf(options.dataLabelPosition || '') < 0)
			options.dataLabelPosition = options.type.name == 'pie' || options.type.name == 'doughnut' ? 'bestFit' : 'ctr'
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
		options.catGridLine = options.catGridLine || (type.name == 'scatter' ? { color: 'D9D9D9', pt: 1 } : 'none')
		options.valGridLine = options.valGridLine || (type.name == 'scatter' ? { color: 'D9D9D9', pt: 1 } : {})
		options.serGridLine = options.serGridLine || (type.name == 'scatter' ? { color: 'D9D9D9', pt: 1 } : 'none')
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
			: options.type.name == 'pie' || options.type.name == 'doughnut'
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
		if (!options.dataLabelFormatCode && options.type.name === 'scatter') options.dataLabelFormatCode = 'General'
		options.dataLabelFormatCode =
			options.dataLabelFormatCode && typeof options.dataLabelFormatCode === 'string'
				? options.dataLabelFormatCode
				: options.type.name == 'pie' || options.type.name == 'doughnut'
				? '0%'
				: '#,##0'
		//
		// Set default format for Scatter chart labels to custom string if not defined
		if (!options.dataLabelFormatScatter && options.type.name === 'scatter') options.dataLabelFormatScatter = 'custom'
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
			type: 'chart',
			globalId: chartId,
			fileName: 'chart' + chartId + '.xml',
			Target: '/ppt/charts/chart' + chartId + '.xml',
		})
		resultObject.chartRid = chartRelId

		target.data.push(resultObject)
		return resultObject
	},

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
	 * @param {Object} slideDef slide definition
	 * @param {Object} target empty slide object that should be updated by the passed definition
	 */
	createSlideObject: function createSlideObject(slideDef, target) {
		// STEP 1: Add background
		if (slideDef.bkgd) {
			gObjPptxGenerators.addBackgroundDefinition(slideDef.bkgd, target)
		}

		// STEP 2: Add all Slide Master objects in the order they were given (Issue#53)
		if (slideDef.objects && Array.isArray(slideDef.objects) && slideDef.objects.length > 0) {
			slideDef.objects.forEach((object, idx) => {
				var key = Object.keys(object)[0]
				if (MASTER_OBJECTS[key] && key == 'chart')
					gObjPptxGenerators.addChartDefinition(CHART_TYPES[(object.chart.type || '').toUpperCase()], object.chart.data, object.chart.opts, target)
				else if (MASTER_OBJECTS[key] && key == 'image') gObjPptxGenerators.addImageDefinition(object[key], target)
				else if (MASTER_OBJECTS[key] && key == 'line') gObjPptxGenerators.addShapeDefinition(gObjPptxShapes.LINE, object[key], target)
				else if (MASTER_OBJECTS[key] && key == 'rect') gObjPptxGenerators.addShapeDefinition(gObjPptxShapes.RECTANGLE, object[key], target)
				else if (MASTER_OBJECTS[key] && key == 'text') gObjPptxGenerators.addTextDefinition(object[key].text, object[key].options, target, false)
				else if (MASTER_OBJECTS[key] && key == 'placeholder') {
					// TODO: 20180820: Check for existing `name`?
					object[key].options.placeholderName = object[key].options.name
					delete object[key].options.name // remap name for earier handling internally
					object[key].options.placeholderType = object[key].options.type
					delete object[key].options.type // remap name for earier handling internally
					object[key].options.placeholderIdx = 100 + idx
					gObjPptxGenerators.addPlaceholderDefinition(object[key].text, object[key].options, target)
				}
			})
		}

		// STEP 3: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
		if (slideDef.slideNumber && typeof slideDef.slideNumber === 'object') {
			target.slideNumberObj = slideDef.slideNumber
		}
	},

	/**
	 * Transforms a slide object to resulting XML string.
	 * @param {ISlide} slideObject slide object created within gObjPptxGenerators.createSlideObject
	 * @return {string} XML string with <p:cSld> as the root
	 */
	slideObjectToXml: function slideObjectToXml(slideObject: ISlide): string {
		var strSlideXml: string = slideObject.name ? '<p:cSld name="' + slideObject.name + '">' : '<p:cSld>'
		var intTableNum: number = 1

		// STEP 1: Add background
		if (slideObject && slideObject.back) {
			strSlideXml += genXmlColorSelection(false, slideObject.back)
		}

		// STEP 2: Add background image (using Strech) (if any)
		if (slideObject && slideObject.bkgdImgRid) {
			// FIXME: We should be doing this in the slideLayout...
			strSlideXml +=
				'<p:bg>' +
				'<p:bgPr><a:blipFill dpi="0" rotWithShape="1">' +
				'<a:blip r:embed="rId' +
				slideObject.bkgdImgRid +
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
		slideObject.data.forEach((slideItemObj, idx) => {
			var x = 0,
				y = 0,
				cx = getSmartParseNumber('75%', 'X', slideObject.layoutObj),
				cy = 0,
				placeholderObj
			var locationAttr = '',
				shapeType = null

			if (slideObject.layoutObj && slideObject.layoutObj.data && slideItemObj.options && slideItemObj.options.placeholder) {
				placeholderObj = slideObject.layoutObj.data.filter(layoutObj => {
					return layoutObj.options.placeholderName == slideItemObj.options.placeholder
				})[0]
			}

			// A: Set option vars
			slideItemObj.options = slideItemObj.options || {}

			if (slideItemObj.options.w || slideItemObj.options.w == 0) slideItemObj.options.cx = slideItemObj.options.w
			if (slideItemObj.options.h || slideItemObj.options.h == 0) slideItemObj.options.cy = slideItemObj.options.h
			//
			if (slideItemObj.options.x || slideItemObj.options.x == 0) x = getSmartParseNumber(slideItemObj.options.x, 'X', slideObject.layoutObj)
			if (slideItemObj.options.y || slideItemObj.options.y == 0) y = getSmartParseNumber(slideItemObj.options.y, 'Y', slideObject.layoutObj)
			if (slideItemObj.options.cx || slideItemObj.options.cx == 0) cx = getSmartParseNumber(slideItemObj.options.cx, 'X', slideObject.layoutObj)
			if (slideItemObj.options.cy || slideItemObj.options.cy == 0) cy = getSmartParseNumber(slideItemObj.options.cy, 'Y', slideObject.layoutObj)

			// If using a placeholder then inherit it's position
			if (placeholderObj) {
				if (placeholderObj.options.x || placeholderObj.options.x == 0) x = getSmartParseNumber(placeholderObj.options.x, 'X', slideObject.layoutObj)
				if (placeholderObj.options.y || placeholderObj.options.y == 0) y = getSmartParseNumber(placeholderObj.options.y, 'Y', slideObject.layoutObj)
				if (placeholderObj.options.cx || placeholderObj.options.cx == 0) cx = getSmartParseNumber(placeholderObj.options.cx, 'X', slideObject.layoutObj)
				if (placeholderObj.options.cy || placeholderObj.options.cy == 0) cy = getSmartParseNumber(placeholderObj.options.cy, 'Y', slideObject.layoutObj)
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
						(intTableNum * slideObject.numb + 1) +
						'" name="Table ' +
						intTableNum * slideObject.numb +
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
						slideObject.relsMedia.filter(rel => {
							return rel.rId == slideItemObj.imageRid
						})[0].extn == 'svg'
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
						var boxW = sizing.w ? getSmartParseNumber(sizing.w, 'X', slideObject.layoutObj) : cx,
							boxH = sizing.h ? getSmartParseNumber(sizing.h, 'Y', slideObject.layoutObj) : cy,
							boxX = getSmartParseNumber(sizing.x || 0, 'X', slideObject.layoutObj),
							boxY = getSmartParseNumber(sizing.y || 0, 'Y', slideObject.layoutObj)

						strSlideXml += gObjPptxGenerators.imageSizingXml[sizing.type]({ w: width, h: height }, { w: boxW, h: boxH, x: boxX, y: boxY })
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
				getSmartParseNumber(slideObject.slideNumberObj.x, 'X', slideObject.layoutObj) +
				'" y="' +
				getSmartParseNumber(slideObject.slideNumberObj.y, 'Y', slideObject.layoutObj) +
				'"/>' +
				'      <a:ext cx="' +
				(slideObject.slideNumberObj.w ? getSmartParseNumber(slideObject.slideNumberObj.w, 'X', slideObject.layoutObj) : 800000) +
				'" cy="' +
				(slideObject.slideNumberObj.h ? getSmartParseNumber(slideObject.slideNumberObj.h, 'Y', slideObject.layoutObj) : 300000) +
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
			strSlideXml +=
				'<a:p><a:fld id="' + SLDNUMFLDID + '" type="slidenum">' + '<a:rPr lang="en-US" smtClean="0"/><a:t></a:t></a:fld>' + '<a:endParaRPr lang="en-US"/></a:p>'
			strSlideXml += '</p:txBody></p:sp>'
		}

		// STEP 6: Close spTree and finalize slide XML
		strSlideXml += '</p:spTree>'
		strSlideXml += '</p:cSld>'

		// LAST: Return
		return strSlideXml
	},

	/**
	 * Transforms slide relations to XML string.
	 * Extra relations that are not dynamic can be passed using the 2nd arg (e.g. theme relation in master file).
	 * These relations use rId series that starts with 1-increased maximum of rIds used for dynamic relations.
	 *
	 * @param {ISlide} slideObject slide object whose relations are being transformed
	 * @param {Object[]} defaultRels array of default relations (such objects expected: { target: <filepath>, type: <schemepath> })
	 * @return {string} complete XML string ready to be saved as a file
	 */
	slideObjectRelationsToXml: function slideObjectRelationsToXml(slideObject: ISlide, defaultRels): string {
		var lastRid = 0 // stores maximum rId used for dynamic relations
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
		strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
		// Add any rels for this Slide (image/audio/video/youtube/chart)
		slideObject.rels.forEach((rel, idx) => {
			lastRid = Math.max(lastRid, rel.rId)
			if (rel.type.toLowerCase().indexOf('image') > -1) {
				strXml +=
					'<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>'
			} else if (rel.type.toLowerCase().indexOf('chart') > -1) {
				strXml +=
					'<Relationship Id="rId' + rel.rId + '" Target="' + rel.Target + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>'
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
	},

	imageSizingXml: {
		cover: (imgSize, boxDim) => {
			var imgRatio = imgSize.h / imgSize.w,
				boxRatio = boxDim.h / boxDim.w,
				isBoxBased = boxRatio > imgRatio,
				width = isBoxBased ? boxDim.h / imgRatio : boxDim.w,
				height = isBoxBased ? boxDim.h : boxDim.w * imgRatio,
				hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)),
				vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))
			return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '" /><a:stretch/>'
		},
		contain: (imgSize, boxDim) => {
			var imgRatio = imgSize.h / imgSize.w,
				boxRatio = boxDim.h / boxDim.w,
				widthBased = boxRatio > imgRatio,
				width = widthBased ? boxDim.w : boxDim.h / imgRatio,
				height = widthBased ? boxDim.w * imgRatio : boxDim.h,
				hzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.w / width)),
				vzPerc = Math.round(1e5 * 0.5 * (1 - boxDim.h / height))
			return '<a:srcRect l="' + hzPerc + '" r="' + hzPerc + '" t="' + vzPerc + '" b="' + vzPerc + '" /><a:stretch/>'
		},
		crop: (imageSize, boxDim) => {
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
	},

	/**
	 * Based on passed data, creates Excel Worksheet that is used as a data source for a chart.
	 * @param {ISlideRelChart} chartObject chart object
	 * @param {jszip} zip file that the resulting XLSX should be added to
	 * @return {Promise} promise of generating the XLSX file
	 */
	createExcelWorksheet: function createExcelWorksheet(chartObject: ISlideRelChart, zip): Promise<any> {
		var data = chartObject.data

		return new Promise((resolve, reject) => {
			var zipExcel = new JSZip()
			var intBubbleCols = (data.length - 1) * 2 + 1 // 1 for "X-Values", then 2 for every Y-Axis

			// A: Add folders
			zipExcel.folder('_rels')
			zipExcel.folder('docProps')
			zipExcel.folder('xl/_rels')
			zipExcel.folder('xl/tables')
			zipExcel.folder('xl/theme')
			zipExcel.folder('xl/worksheets')
			zipExcel.folder('xl/worksheets/_rels')

			// B: Add core contents
			{
				zipExcel.file(
					'[Content_Types].xml',
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
						'  <Default Extension="xml" ContentType="application/xml"/>' +
						'  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
						//+ '  <Default Extension="jpeg" ContentType="image/jpg"/><Default Extension="png" ContentType="image/png"/>'
						//+ '  <Default Extension="bmp" ContentType="image/bmp"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="tif" ContentType="image/tif"/><Default Extension="pdf" ContentType="application/pdf"/><Default Extension="mov" ContentType="application/movie"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
						//+ '  <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
						'  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
						'  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
						'  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
						'  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
						'  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
						'  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
						'  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
						'  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
						'</Types>\n'
				)
				zipExcel.file(
					'_rels/.rels',
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
						'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
						'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
						'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
						'</Relationships>\n'
				)
				zipExcel.file(
					'docProps/app.xml',
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
						'<Application>Microsoft Excel</Application>' +
						'<DocSecurity>0</DocSecurity>' +
						'<ScaleCrop>false</ScaleCrop>' +
						'<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>' +
						'</Properties>\n'
				)
				zipExcel.file(
					'docProps/core.xml',
					'<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
						'<dc:creator>PptxGenJS</dc:creator>' +
						'<cp:lastModifiedBy>Ely, Brent</cp:lastModifiedBy>' +
						'<dcterms:created xsi:type="dcterms:W3CDTF">' +
						new Date().toISOString() +
						'</dcterms:created>' +
						'<dcterms:modified xsi:type="dcterms:W3CDTF">' +
						new Date().toISOString() +
						'</dcterms:modified>' +
						'</cp:coreProperties>\n'
				)
				zipExcel.file(
					'xl/_rels/workbook.xml.rels',
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
						'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
						'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
						'<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
						'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
						'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
						'</Relationships>\n'
				)
				zipExcel.file(
					'xl/styles.xml',
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/>' +
						'<name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/>' +
						'<rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n'
				)
				zipExcel.file(
					'xl/theme/theme1.xml',
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>'
				)
				zipExcel.file(
					'xl/workbook.xml',
					'<?xml version="1.0" encoding="UTF-8"?>' +
						'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
						'<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/>' +
						'<workbookPr />' +
						'<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="15960" windowHeight="18080"/></bookViews>' +
						'<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" /></sheets>' +
						'<calcPr calcId="171026" concurrentCalc="0"/>' +
						'</workbook>\n'
				)
				zipExcel.file(
					'xl/worksheets/_rels/sheet1.xml.rels',
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
						'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
						'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
						'</Relationships>\n'
				)
			}

			// sharedStrings.xml
			{
				// A: Start XML
				var strSharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
				if (chartObject.opts.type === 'bubble') {
					strSharedStrings +=
						'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (intBubbleCols + 1) + '" uniqueCount="' + (intBubbleCols + 1) + '">'
				} else if (chartObject.opts.type === 'scatter') {
					strSharedStrings +=
						'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (data.length + 1) + '" uniqueCount="' + (data.length + 1) + '">'
				} else {
					strSharedStrings +=
						'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' +
						(data[0].labels.length + data.length + 1) +
						'" uniqueCount="' +
						(data[0].labels.length + data.length + 1) +
						'">'
					// B: Add 'blank' for A1
					strSharedStrings += '<si><t xml:space="preserve"></t></si>'
				}

				// C: Add `name`/Series
				if (chartObject.opts.type === 'bubble') {
					data.forEach((objData, idx) => {
						if (idx == 0) strSharedStrings += '<si><t>' + 'X-Axis' + '</t></si>'
						else {
							strSharedStrings += '<si><t>' + encodeXmlEntities(objData.name || ' ') + '</t></si>'
							strSharedStrings += '<si><t>' + encodeXmlEntities('Size ' + idx) + '</t></si>'
						}
					})
				} else {
					data.forEach(objData => {
						strSharedStrings += '<si><t>' + encodeXmlEntities((objData.name || ' ').replace('X-Axis', 'X-Values')) + '</t></si>'
					})
				}

				// D: Add `labels`/Categories
				if (chartObject.opts.type != 'bubble' && chartObject.opts.type != 'scatter') {
					data[0].labels.forEach(label => {
						strSharedStrings += '<si><t>' + encodeXmlEntities(label) + '</t></si>'
					})
				}

				strSharedStrings += '</sst>\n'
				zipExcel.file('xl/sharedStrings.xml', strSharedStrings)
			}

			// tables/table1.xml
			{
				var strTableXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
				if (chartObject.opts.type == 'bubble') {
					/*
					strTableXml += '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:'+ LETTERS[data.length-1] + (data[0].values.length+1) +'" totalsRowShown="0">';
					strTableXml += '<tableColumns count="' + (data.length) +'">';
					data.forEach(function(obj,idx){ strTableXml += '<tableColumn id="'+ (idx+1) +'" name="'+ (idx==0 ? 'X-Values' : 'Y-Value '+idx) +'" />' });
					*/
				} else if (chartObject.opts.type == 'scatter') {
					strTableXml +=
						'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' +
						LETTERS[data.length - 1] +
						(data[0].values.length + 1) +
						'" totalsRowShown="0">'
					strTableXml += '<tableColumns count="' + data.length + '">'
					data.forEach((_obj, idx) => {
						strTableXml += '<tableColumn id="' + (idx + 1) + '" name="' + (idx == 0 ? 'X-Values' : 'Y-Value ' + idx) + '" />'
					})
				} else {
					strTableXml +=
						'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' +
						LETTERS[data.length] +
						(data[0].labels.length + 1) +
						'" totalsRowShown="0">'
					strTableXml += '<tableColumns count="' + (data.length + 1) + '">'
					strTableXml += '<tableColumn id="1" name=" " />'
					data.forEach((obj, idx) => {
						strTableXml += '<tableColumn id="' + (idx + 2) + '" name="' + encodeXmlEntities(obj.name) + '" />'
					})
				}
				strTableXml += '</tableColumns>'
				strTableXml += '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0" />'
				strTableXml += '</table>'
				zipExcel.file('xl/tables/table1.xml', strTableXml)
			}

			// worksheets/sheet1.xml
			{
				var strSheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
				strSheetXml +=
					'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
				if (chartObject.opts.type === 'bubble') {
					strSheetXml += '<dimension ref="A1:' + LETTERS[intBubbleCols - 1] + (data[0].values.length + 1) + '" />'
				} else if (chartObject.opts.type === 'scatter') {
					strSheetXml += '<dimension ref="A1:' + LETTERS[data.length - 1] + (data[0].values.length + 1) + '" />'
				} else {
					strSheetXml += '<dimension ref="A1:' + LETTERS[data.length] + (data[0].labels.length + 1) + '" />'
				}

				strSheetXml += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1" /></sheetView></sheetViews>'
				strSheetXml += '<sheetFormatPr baseColWidth="10" defaultColWidth="11.5" defaultRowHeight="12" />'
				if (chartObject.opts.type == 'bubble') {
					strSheetXml += '<cols>'
					strSheetXml += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />'
					strSheetXml += '</cols>'
					/* EX: INPUT: `data`
					[
						{ name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
						{ name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9], sizes:[ 4, 5, 6, 7, 8] },
						{ name:'Y-Axis 2', values:[33,32,42,53,63], sizes:[11,12,13,14,15] }
					];
					*/
					/* EX: OUTPUT: bubbleChart Worksheet:
						-|----A-----|------B-----|------C-----|------D-----|------E-----|
						1| X-Values | Y-Values 1 | Y-Sizes 1  | Y-Values 2 | Y-Sizes 2  |
						2|    11    |     22     |      4     |     33     |      8     |
						-|----------|------------|------------|------------|------------|
					*/
					strSheetXml += '<sheetData>'

					// A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
					strSheetXml += '<row r="1" spans="1:' + intBubbleCols + '">'
					strSheetXml += '<c r="A1" t="s"><v>0</v></c>'
					for (var idx = 1; idx < intBubbleCols; idx++) {
						strSheetXml += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">' // NOTE: use `t="s"` for label cols!
						strSheetXml += '<v>' + idx + '</v>'
						strSheetXml += '</c>'
					}
					strSheetXml += '</row>'

					// B: Add row for each X-Axis value (Y-Axis* value is optional)
					data[0].values.forEach((val, idx) => {
						// Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
						strSheetXml += '<row r="' + (idx + 2) + '" spans="1:' + intBubbleCols + '">'
						strSheetXml += '<c r="A' + (idx + 2) + '"><v>' + val + '</v></c>'
						// Add Y-Axis 1->N (idy=0 = Xaxis)
						var idxColLtr = 1
						for (var idy = 1; idy < data.length; idy++) {
							// y-value
							strSheetXml += '<c r="' + (idxColLtr < 26 ? LETTERS[idxColLtr] : 'A' + LETTERS[idxColLtr % LETTERS.length]) + '' + (idx + 2) + '">'
							strSheetXml += '<v>' + (data[idy].values[idx] || '') + '</v>'
							strSheetXml += '</c>'
							idxColLtr++
							// y-size
							strSheetXml += '<c r="' + (idxColLtr < 26 ? LETTERS[idxColLtr] : 'A' + LETTERS[idxColLtr % LETTERS.length]) + '' + (idx + 2) + '">'
							strSheetXml += '<v>' + (data[idy].sizes[idx] || '') + '</v>'
							strSheetXml += '</c>'
							idxColLtr++
						}
						strSheetXml += '</row>'
					})
				} else if (chartObject.opts.type == 'scatter') {
					strSheetXml += '<cols>'
					strSheetXml += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />'
					//data.forEach((obj,idx)=>{ strSheetXml += '<col min="'+(idx+1)+'" max="'+(idx+1)+'" width="11" customWidth="1" />' });
					strSheetXml += '</cols>'
					/* EX: INPUT: `data`
					[
						{ name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
						{ name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9] },
						{ name:'Y-Axis 2', values:[33,32,42,53,63] }
					];
					*/
					/* EX: OUTPUT: scatterChart Worksheet:
						-|----A-----|------B-----|
						1| X-Values | Y-Values 1 |
						2|    11    |     22     |
						-|----------|------------|
					*/
					strSheetXml += '<sheetData>'

					// A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
					strSheetXml += '<row r="1" spans="1:' + data.length + '">'
					strSheetXml += '<c r="A1" t="s"><v>0</v></c>'
					for (var idx = 1; idx < data.length; idx++) {
						strSheetXml += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">' // NOTE: use `t="s"` for label cols!
						strSheetXml += '<v>' + idx + '</v>'
						strSheetXml += '</c>'
					}
					strSheetXml += '</row>'

					// B: Add row for each X-Axis value (Y-Axis* value is optional)
					data[0].values.forEach((val, idx) => {
						// Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
						strSheetXml += '<row r="' + (idx + 2) + '" spans="1:' + data.length + '">'
						strSheetXml += '<c r="A' + (idx + 2) + '"><v>' + val + '</v></c>'
						// Add Y-Axis 1->N
						for (var idy = 1; idy < data.length; idy++) {
							strSheetXml += '<c r="' + (idy < 26 ? LETTERS[idy] : 'A' + LETTERS[idy % LETTERS.length]) + '' + (idx + 2) + '">'
							strSheetXml += '<v>' + (data[idy].values[idx] || data[idy].values[idx] == 0 ? data[idy].values[idx] : '') + '</v>'
							strSheetXml += '</c>'
						}
						strSheetXml += '</row>'
					})
				} else {
					strSheetXml += '<cols>'
					strSheetXml += '<col min="1" max="1" width="11" customWidth="1" />'
					//data.forEach(function(){ strSheetXml += '<col min="10" max="100" width="10" customWidth="1" />' });
					strSheetXml += '</cols>'
					strSheetXml += '<sheetData>'

					/* EX: INPUT: `data`
					[
						{ name:'Red', labels:['Jan..May-17'], values:[11,13,14,15,16] },
						{ name:'Amb', labels:['Jan..May-17'], values:[22, 6, 7, 8, 9] },
						{ name:'Grn', labels:['Jan..May-17'], values:[33,32,42,53,63] }
					];
					*/
					/* EX: OUTPUT: lineChart Worksheet:
						-|---A---|--B--|--C--|--D--|
						1|       | Red | Amb | Grn |
						2|Jan-17 |   11|   22|   33|
						3|Feb-17 |   55|   43|   70|
						4|Mar-17 |   56|  143|   99|
						5|Apr-17 |   65|    3|  120|
						6|May-17 |   75|   93|  170|
						-|-------|-----|-----|-----|
					*/

					// A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
					strSheetXml += '<row r="1" spans="1:' + (data.length + 1) + '">'
					strSheetXml += '<c r="A1" t="s"><v>0</v></c>'
					for (var idx = 1; idx <= data.length; idx++) {
						// FIXME: Max cols is 52
						strSheetXml += '<c r="' + (idx < 26 ? LETTERS[idx] : 'A' + LETTERS[idx % LETTERS.length]) + '1" t="s">' // NOTE: use `t="s"` for label cols!
						strSheetXml += '<v>' + idx + '</v>'
						strSheetXml += '</c>'
					}
					strSheetXml += '</row>'

					// B: Add data row(s) for each category
					data[0].labels.forEach((cat, idx) => {
						// Leading col is reserved for the label, so hard-code it, then loop over col values
						strSheetXml += '<row r="' + (idx + 2) + '" spans="1:' + (data.length + 1) + '">'
						strSheetXml += '<c r="A' + (idx + 2) + '" t="s">'
						strSheetXml += '<v>' + (data.length + idx + 1) + '</v>'
						strSheetXml += '</c>'
						for (var idy = 0; idy < data.length; idy++) {
							strSheetXml += '<c r="' + (idy + 1 < 26 ? LETTERS[idy + 1] : 'A' + LETTERS[(idy + 1) % LETTERS.length]) + '' + (idx + 2) + '">'
							strSheetXml += '<v>' + (data[idy].values[idx] || '') + '</v>'
							strSheetXml += '</c>'
						}
						strSheetXml += '</row>'
					})
				}
				strSheetXml += '</sheetData>'
				strSheetXml += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />'
				// Link the `table1.xml` file to define an actual Table in Excel
				// NOTE: This onyl works with scatter charts - all others give a "cannot find linked file" error
				// ....: Since we dont need the table anyway (chart data can be edited/range selected, etc.), just dont use this
				// ....: Leaving this so nobody foolishly attempts to add this in the future
				// strSheetXml += '<tableParts count="1"><tablePart r:id="rId1" /></tableParts>';
				strSheetXml += '</worksheet>\n'
				zipExcel.file('xl/worksheets/sheet1.xml', strSheetXml)
			}

			// C: Add XLSX to PPTX export
			zipExcel
				.generateAsync({ type: 'base64' })
				.then(content => {
					// 1: Create the embedded Excel worksheet with labels and data
					zip.file('ppt/embeddings/Microsoft_Excel_Worksheet' + chartObject.globalId + '.xlsx', content, { base64: true })

					// 2: Create the chart.xml and rels files
					zip.file(
						'ppt/charts/_rels/' + chartObject.fileName + '.rels',
						'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
							'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
							'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet' +
							chartObject.globalId +
							'.xlsx"/>' +
							'</Relationships>'
					)
					zip.file('ppt/charts/' + chartObject.fileName, makeXmlCharts(chartObject))

					// 3: Done
					resolve()
				})
				.catch(strErr => {
					reject(strErr)
				})
		})
	},
}

/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 */
export function makeXmlCharts(rel: ISlideRelChart) {
	/* ----------------------------------------------------------------------- */

	// STEP 1: Create chart
	{
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		// CHARTSPACE: BEGIN vvv
		strXml +=
			'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
		strXml += '<c:date1904 val="0"/>' // ppt defaults to 1904 dates, excel to 1900
		strXml += '<c:chart>'

		// OPTION: Title
		if (rel.opts.showTitle) {
			strXml += genXmlTitle({
				title: rel.opts.title || 'Chart Title',
				fontSize: rel.opts.titleFontSize || DEF_FONT_TITLE_SIZE,
				color: rel.opts.titleColor,
				fontFace: rel.opts.titleFontFace,
				rotate: rel.opts.titleRotate,
				titleAlign: rel.opts.titleAlign,
				titlePos: rel.opts.titlePos,
			})
			strXml += '<c:autoTitleDeleted val="0"/>'
		} else {
			// NOTE: Add autoTitleDeleted tag in else to prevent default creation of chart title even when showTitle is set to false
			strXml += '<c:autoTitleDeleted val="1"/>'
		}
		// Add 3D view tag
		if (rel.opts.type == 'bar3D') {
			strXml += '<c:view3D>'
			strXml += ' <c:rotX val="' + rel.opts.v3DRotX + '"/>'
			strXml += ' <c:rotY val="' + rel.opts.v3DRotY + '"/>'
			strXml += ' <c:rAngAx val="' + (rel.opts.v3DRAngAx == false ? 0 : 1) + '"/>'
			strXml += ' <c:perspective val="' + rel.opts.v3DPerspective + '"/>'
			strXml += '</c:view3D>'
		}

		strXml += '<c:plotArea>'
		// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
		if (rel.opts.layout) {
			strXml += '<c:layout>'
			strXml += ' <c:manualLayout>'
			strXml += '  <c:layoutTarget val="inner" />'
			strXml += '  <c:xMode val="edge" />'
			strXml += '  <c:yMode val="edge" />'
			strXml += '  <c:x val="' + (rel.opts.layout.x || 0) + '" />'
			strXml += '  <c:y val="' + (rel.opts.layout.y || 0) + '" />'
			strXml += '  <c:w val="' + (rel.opts.layout.w || 1) + '" />'
			strXml += '  <c:h val="' + (rel.opts.layout.h || 1) + '" />'
			strXml += ' </c:manualLayout>'
			strXml += '</c:layout>'
		} else {
			strXml += '<c:layout/>'
		}
	}

	var usesSecondaryValAxis = false

	// A: Create Chart XML -----------------------------------------------------------
	if (Array.isArray(rel.opts.type)) {
		rel.opts.type.forEach((type: ISlideRelChart) => {
			///var options = getMix(rel.opts, type.options) as IChartOpts; // TODO: FIXME: theres `options` on chart rels??
			let options: IChartOpts = {
				type: type.type,
			}
			var valAxisId = options['secondaryValAxis'] ? AXIS_ID_VALUE_SECONDARY : AXIS_ID_VALUE_PRIMARY
			var catAxisId = options['secondaryCatAxis'] ? AXIS_ID_CATEGORY_SECONDARY : AXIS_ID_CATEGORY_PRIMARY
			usesSecondaryValAxis = usesSecondaryValAxis || options['secondaryValAxis']
			strXml += makeChartType(type.type, type.data, options, valAxisId, catAxisId, true)
		})
	} else {
		strXml += makeChartType(rel.opts.type, rel.data, rel.opts, AXIS_ID_VALUE_PRIMARY, AXIS_ID_CATEGORY_PRIMARY, false)
	}

	// B: Axes -----------------------------------------------------------
	if (rel.opts.type !== 'pie' && rel.opts.type !== 'doughnut') {
		// Param check
		if (rel.opts.valAxes && !usesSecondaryValAxis) {
			throw new Error('Secondary axis must be used by one of the multiple charts')
		}

		if (rel.opts.catAxes) {
			if (!rel.opts.valAxes || rel.opts.valAxes.length !== rel.opts.catAxes.length) {
				throw new Error('There must be the same number of value and category axes.')
			}
			strXml += makeCatAxis(getMix(rel.opts, rel.opts.catAxes[0]), AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY)
			if (rel.opts.catAxes[1]) {
				strXml += makeCatAxis(getMix(rel.opts, rel.opts.catAxes[1]), AXIS_ID_CATEGORY_SECONDARY, AXIS_ID_VALUE_PRIMARY)
			}
		} else {
			strXml += makeCatAxis(rel.opts, AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY)
		}

		if (rel.opts.valAxes) {
			strXml += makeValueAxis(getMix(rel.opts, rel.opts.valAxes[0]) as IChartOpts, AXIS_ID_VALUE_PRIMARY)
			if (rel.opts.valAxes[1]) {
				strXml += makeValueAxis(getMix(rel.opts, rel.opts.valAxes[1]) as IChartOpts, AXIS_ID_VALUE_SECONDARY)
			}
		} else {
			strXml += makeValueAxis(rel.opts, AXIS_ID_VALUE_PRIMARY)

			// Add series axis for 3D bar
			if (rel.opts.type == 'bar3D') {
				strXml += makeSerAxis(rel.opts, AXIS_ID_SERIES_PRIMARY, AXIS_ID_VALUE_PRIMARY)
			}
		}
	}

	// C: Chart Properties and plotArea Options: Border, Data Table, Fill, Legend
	{
		// NOTE: DataTable goes between '</c:valAx>' and '<c:spPr>'
		if (rel.opts.showDataTable) {
			strXml += '<c:dTable>'
			strXml += '  <c:showHorzBorder val="' + (rel.opts.showDataTableHorzBorder == false ? 0 : 1) + '"/>'
			strXml += '  <c:showVertBorder val="' + (rel.opts.showDataTableVertBorder == false ? 0 : 1) + '"/>'
			strXml += '  <c:showOutline    val="' + (rel.opts.showDataTableOutline == false ? 0 : 1) + '"/>'
			strXml += '  <c:showKeys       val="' + (rel.opts.showDataTableKeys == false ? 0 : 1) + '"/>'
			strXml += '  <c:spPr>'
			strXml += '    <a:noFill/>'
			strXml +=
				'    <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>'
			strXml += '    <a:effectLst/>'
			strXml += '  </c:spPr>'
			strXml +=
				'  <c:txPr>\
						  <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
						  <a:lstStyle/>\
						  <a:p>\
							<a:pPr rtl="0">\
							  <a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
								<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>\
								<a:latin typeface="+mn-lt"/>\
								<a:ea typeface="+mn-ea"/>\
								<a:cs typeface="+mn-cs"/>\
							  </a:defRPr>\
							</a:pPr>\
							<a:endParaRPr lang="en-US"/>\
						  </a:p>\
						</c:txPr>\
					  </c:dTable>'
		}

		strXml += '  <c:spPr>'

		// OPTION: Fill
		strXml += rel.opts.fill ? genXmlColorSelection(rel.opts.fill) : '<a:noFill/>'

		// OPTION: Border
		strXml += rel.opts.border
			? '<a:ln w="' + rel.opts.border.pt * ONEPT + '"' + ' cap="flat">' + genXmlColorSelection(rel.opts.border.color) + '</a:ln>'
			: '<a:ln><a:noFill/></a:ln>'

		// Close shapeProp/plotArea before Legend
		strXml += '    <a:effectLst/>'
		strXml += '  </c:spPr>'
		strXml += '</c:plotArea>'

		// OPTION: Legend
		// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
		if (rel.opts.showLegend) {
			strXml += '<c:legend>'
			strXml += '<c:legendPos val="' + rel.opts.legendPos + '"/>'
			strXml += '<c:layout/>'
			strXml += '<c:overlay val="0"/>'
			if (rel.opts.legendFontFace || rel.opts.legendFontSize || rel.opts.legendColor) {
				strXml += '<c:txPr>'
				strXml += '  <a:bodyPr/>'
				strXml += '  <a:lstStyle/>'
				strXml += '  <a:p>'
				strXml += '    <a:pPr>'
				strXml += rel.opts.legendFontSize ? '<a:defRPr sz="' + Number(rel.opts.legendFontSize) * 100 + '">' : '<a:defRPr>'
				if (rel.opts.legendColor) strXml += genXmlColorSelection(rel.opts.legendColor)
				if (rel.opts.legendFontFace) strXml += '<a:latin typeface="' + rel.opts.legendFontFace + '"/>'
				if (rel.opts.legendFontFace) strXml += '<a:cs    typeface="' + rel.opts.legendFontFace + '"/>'
				strXml += '      </a:defRPr>'
				strXml += '    </a:pPr>'
				strXml += '    <a:endParaRPr lang="en-US"/>'
				strXml += '  </a:p>'
				strXml += '</c:txPr>'
			}
			strXml += '</c:legend>'
		}
	}

	strXml += '  <c:plotVisOnly val="1"/>'
	strXml += '  <c:dispBlanksAs val="' + rel.opts.displayBlanksAs + '"/>'
	if (rel.opts.type === 'scatter') strXml += '<c:showDLblsOverMax val="1"/>'

	strXml += '</c:chart>'

	// D: CHARTSPACE SHAPE PROPS
	strXml += '<c:spPr>'
	strXml += '  <a:noFill/>'
	strXml += '  <a:ln w="12700" cap="flat"><a:noFill/><a:miter lim="400000"/></a:ln>'
	strXml += '  <a:effectLst/>'
	strXml += '</c:spPr>'

	// E: DATA (Add relID)
	strXml += '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'

	// LAST: chartSpace end
	strXml += '</c:chartSpace>'

	return strXml
}

/**
 * Create XML string for any given chart type
 * @example: <c:bubbleChart> or <c:lineChart>
 *
 * @param {String} CHART_TYPES key
 * @param {String} data
 * @param {object} opts
 * @param {String} valAxisId
 * @param {String} catAxisId
 * @param {boolean} isMultiTypeChart
 */
function makeChartType(chartType: ISlideRelChart['type'], data: ISlideRelChart['data'], opts: IChartOpts, valAxisId: string, catAxisId: string, isMultiTypeChart: boolean) {
	// NOTE: "Chart Range" (as shown in "select Chart Area dialog") is calculated.
	// ....: Ensure each X/Y Axis/Col has same row height (esp. applicable to XY Scatter where X can often be larger than Y's)
	var strXml: string = ''

	switch (chartType) {
		case 'area':
		case 'bar':
		case 'bar3D':
		case 'line':
		case 'radar':
			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>'
			if (chartType == 'bar' || chartType == 'bar3D') {
				strXml += '<c:barDir val="' + opts.barDir + '"/>'
				strXml += '<c:grouping val="' + opts.barGrouping + '"/>'
			}

			if (chartType == 'radar') {
				strXml += '<c:radarStyle val="' + opts.radarStyle + '"/>'
			}

			strXml += '<c:varyColors val="0"/>'

			// 2: "Series" block for every data row
			/* EX:
				data: [
				 {
				   name: 'Region 1',
				   labels: ['April', 'May', 'June', 'July'],
				   values: [17, 26, 53, 96]
				 },
				 {
				   name: 'Region 2',
				   labels: ['April', 'May', 'June', 'July'],
				   values: [55, 43, 70, 58]
				 }
				]
			*/
			var colorIndex = -1 // Maintain the color index by region
			data.forEach(obj => {
				colorIndex++
				var idx = obj.index
				strXml += '<c:ser>'
				strXml += '  <c:idx val="' + idx + '"/>'
				strXml += '  <c:order val="' + idx + '"/>'
				strXml += '  <c:tx>'
				strXml += '    <c:strRef>'
				strXml += '      <c:f>Sheet1!$' + getExcelColName(idx + 1) + '$1</c:f>'
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt></c:strCache>'
				strXml += '    </c:strRef>'
				strXml += '  </c:tx>'
				strXml += '  <c:invertIfNegative val="0"/>'

				// Fill and Border
				var strSerColor = opts.chartColors[colorIndex % opts.chartColors.length]

				strXml += '  <c:spPr>'
				if (strSerColor == 'transparent') {
					strXml += '<a:noFill/>'
				} else if (opts.chartColorsOpacity) {
					strXml += '<a:solidFill>' + createColorElement(strSerColor, '<a:alpha val="' + opts.chartColorsOpacity + '000"/>') + '</a:solidFill>'
				} else {
					strXml += '<a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>'
				}

				if (chartType == 'line') {
					if (opts.lineSize == 0) {
						strXml += '<a:ln><a:noFill/></a:ln>'
					} else {
						strXml += '<a:ln w="' + opts.lineSize * ONEPT + '" cap="flat"><a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>'
						strXml += '<a:prstDash val="' + (opts.lineDash || 'solid') + '"/><a:round/></a:ln>'
					}
				} else if (opts.dataBorder) {
					strXml +=
						'<a:ln w="' +
						opts.dataBorder.pt * ONEPT +
						'" cap="flat"><a:solidFill>' +
						createColorElement(opts.dataBorder.color) +
						'</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'
				}

				strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW)

				strXml += '  </c:spPr>'

				// Data Labels per series
				// [20190117] NOTE: Adding these to RADAR chart causes unrecoverable corruption!
				if (chartType != 'radar') {
					strXml += '  <c:dLbls>'
					strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>'
					if (opts.dataLabelBkgrdColors) {
						strXml += '    <c:spPr>'
						strXml += '       <a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>'
						strXml += '    </c:spPr>'
					}
					strXml += '    <c:txPr>'
					strXml += '      <a:bodyPr/>'
					strXml += '      <a:lstStyle/>'
					strXml += '      <a:p><a:pPr>'
					strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">'
					strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>'
					strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>'
					strXml += '        </a:defRPr>'
					strXml += '      </a:pPr></a:p>'
					strXml += '    </c:txPr>'
					// Setting dLblPos tag for bar3D seems to break the generated chart
					if (chartType != 'area' && chartType != 'bar3D') {
						strXml += '<c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>'
					}
					strXml += '    <c:showLegendKey val="0"/>'
					strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
					strXml += '    <c:showCatName val="0"/>'
					strXml += '    <c:showSerName val="0"/>'
					strXml += '    <c:showPercent val="0"/>'
					strXml += '    <c:showBubbleSize val="0"/>'
					strXml += '    <c:showLeaderLines val="0"/>'
					strXml += '  </c:dLbls>'
				}

				// 'c:marker' tag: `lineDataSymbol`
				if (chartType == 'line' || chartType == 'radar') {
					strXml += '<c:marker>'
					strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>'
					if (opts.lineDataSymbolSize) {
						// Defaults to "auto" otherwise (but this is usually too small, so there is a default)
						strXml += '  <c:size val="' + opts.lineDataSymbolSize + '"/>'
					}
					strXml += '  <c:spPr>'
					strXml +=
						'    <a:solidFill>' +
						createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]) +
						'</a:solidFill>'

					var symbolLineColor = opts.lineDataSymbolLineColor || strSerColor
					strXml +=
						'    <a:ln w="' +
						opts.lineDataSymbolLineSize +
						'" cap="flat"><a:solidFill>' +
						createColorElement(symbolLineColor) +
						'</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'
					strXml += '    <a:effectLst/>'
					strXml += '  </c:spPr>'
					strXml += '</c:marker>'
				}

				// Color chart bars various colors
				// Allow users with a single data set to pass their own array of colors (check for this using != ours)
				if ((chartType == 'bar' || chartType == 'bar3D') && (data.length === 1 || opts.valueBarColors) && opts.chartColors != BARCHART_COLORS) {
					// Series Data Point colors
					obj.values.forEach((value, index) => {
						var arrColors = value < 0 ? opts.invertedColors || BARCHART_COLORS : opts.chartColors

						strXml += '  <c:dPt>'
						strXml += '    <c:idx val="' + index + '"/>'
						strXml += '      <c:invertIfNegative val="' + (opts.invertedColors ? 0 : 1) + '"/>'
						strXml += '    <c:bubble3D val="0"/>'
						strXml += '    <c:spPr>'
						if (opts.lineSize === 0) {
							strXml += '<a:ln><a:noFill/></a:ln>'
						} else if (chartType === 'bar') {
							strXml += '<a:solidFill>'
							strXml += '  <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>'
							strXml += '</a:solidFill>'
						} else {
							strXml += '<a:ln>'
							strXml += '  <a:solidFill>'
							strXml += '   <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>'
							strXml += '  </a:solidFill>'
							strXml += '</a:ln>'
						}
						strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW)
						strXml += '    </c:spPr>'
						strXml += '  </c:dPt>'
					})
				}

				// 2: "Categories"
				{
					strXml += '<c:cat>'
					if (opts.catLabelFormatCode) {
						// Use 'numRef' as catLabelFormatCode implies that we are expecting numbers here
						strXml += '  <c:numRef>'
						strXml += '    <c:f>Sheet1!' + '$A$2:$A$' + (obj.labels.length + 1) + '</c:f>'
						strXml += '    <c:numCache>'
						strXml += '      <c:formatCode>' + opts.catLabelFormatCode + '</c:formatCode>'
						strXml += '      <c:ptCount val="' + obj.labels.length + '"/>'
						obj.labels.forEach((label, idx) => {
							strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>'
						})
						strXml += '    </c:numCache>'
						strXml += '  </c:numRef>'
					} else {
						strXml += '  <c:strRef>'
						strXml += '    <c:f>Sheet1!' + '$A$2:$A$' + (obj.labels.length + 1) + '</c:f>'
						strXml += '    <c:strCache>'
						strXml += '	     <c:ptCount val="' + obj.labels.length + '"/>'
						obj.labels.forEach((label, idx) => {
							strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>'
						})
						strXml += '    </c:strCache>'
						strXml += '  </c:strRef>'
					}
					strXml += '</c:cat>'
				}

				// 3: "Values"
				{
					strXml += '  <c:val>'
					strXml += '    <c:numRef>'
					strXml += '      <c:f>Sheet1!' + '$' + getExcelColName(idx + 1) + '$2:$' + getExcelColName(idx + 1) + '$' + (obj.labels.length + 1) + '</c:f>'
					strXml += '      <c:numCache>'
					strXml += '        <c:formatCode>General</c:formatCode>'
					strXml += '	       <c:ptCount val="' + obj.labels.length + '"/>'
					obj.values.forEach((value, idx) => {
						strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value == 0 ? value : '') + '</c:v></c:pt>'
					})
					strXml += '      </c:numCache>'
					strXml += '    </c:numRef>'
					strXml += '  </c:val>'
				}

				// Option: `smooth`
				if (chartType == 'line') strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>'

				// 4: Close "SERIES"
				strXml += '</c:ser>'
			})

			// 3: "Data Labels"
			{
				strXml += '  <c:dLbls>'
				strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>'
				strXml += '    <c:txPr>'
				strXml += '      <a:bodyPr/>'
				strXml += '      <a:lstStyle/>'
				strXml += '      <a:p><a:pPr>'
				strXml +=
					'        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">'
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>'
				strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>'
				strXml += '        </a:defRPr>'
				strXml += '      </a:pPr></a:p>'
				strXml += '    </c:txPr>'
				// NOTE: Throwing an error while creating a multi type chart which contains area chart as the below line appears for the other chart type.
				// Either the given change can be made or the below line can be removed to stop the slide containing multi type chart with area to crash.
				if (opts.type != 'area' && opts.type != 'radar' && !isMultiTypeChart) strXml += '<c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>'
				strXml += '    <c:showLegendKey val="0"/>'
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
				strXml += '    <c:showCatName val="0"/>'
				strXml += '    <c:showSerName val="0"/>'
				strXml += '    <c:showPercent val="0"/>'
				strXml += '    <c:showBubbleSize val="0"/>'
				strXml += '    <c:showLeaderLines val="0"/>'
				strXml += '  </c:dLbls>'
			}

			// 4: Add more chart options (gapWidth, line Marker, etc.)
			if (chartType == 'bar') {
				strXml += '  <c:gapWidth val="' + opts.barGapWidthPct + '"/>'
				strXml += '  <c:overlap val="' + (opts.barGrouping.indexOf('tacked') > -1 ? 100 : 0) + '"/>'
			} else if (chartType == 'bar3D') {
				strXml += '  <c:gapWidth val="' + opts.barGapWidthPct + '"/>'
				strXml += '  <c:gapDepth val="' + opts.barGapDepthPct + '"/>'
				strXml += '  <c:shape val="' + opts.bar3DShape + '"/>'
			} else if (chartType == 'line') {
				strXml += '  <c:marker val="1"/>'
			}

			// 5: Add axisId (NOTE: order matters! (category comes first))
			strXml += '  <c:axId val="' + catAxisId + '"/>'
			strXml += '  <c:axId val="' + valAxisId + '"/>'
			strXml += '  <c:axId val="' + AXIS_ID_SERIES_PRIMARY + '"/>'

			// 6: Close Chart tag
			strXml += '</c:' + chartType + 'Chart>'

			// end switch
			break

		case 'scatter':
			/*
				`data` = [
					{ name:'X-Axis',    values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Value 1', values:[13, 20, 21, 25] },
					{ name:'Y-Value 2', values:[ 1,  2,  5,  9] }
				];
			*/

			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>'
			strXml += '<c:scatterStyle val="lineMarker"/>'
			strXml += '<c:varyColors val="0"/>'

			// 2: Series: (One for each Y-Axis)
			var colorIndex = -1
			data.filter((_obj, idx) => {
				return idx > 0
			}).forEach((obj, idx) => {
				colorIndex++
				strXml += '<c:ser>'
				strXml += '  <c:idx val="' + idx + '"/>'
				strXml += '  <c:order val="' + idx + '"/>'
				strXml += '  <c:tx>'
				strXml += '    <c:strRef>'
				strXml += '      <c:f>Sheet1!$' + LETTERS[idx + 1] + '$1</c:f>'
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + obj.name + '</c:v></c:pt></c:strCache>'
				strXml += '    </c:strRef>'
				strXml += '  </c:tx>'

				// 'c:spPr': Fill, Border, Line, LineStyle (dash, etc.), Shadow
				strXml += '  <c:spPr>'
				{
					var strSerColor = opts.chartColors[colorIndex % opts.chartColors.length]

					if (strSerColor == 'transparent') {
						strXml += '<a:noFill/>'
					} else if (opts.chartColorsOpacity) {
						strXml += '<a:solidFill>' + createColorElement(strSerColor, '<a:alpha val="' + opts.chartColorsOpacity + '000"/>') + '</a:solidFill>'
					} else {
						strXml += '<a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>'
					}

					if (opts.lineSize == 0) {
						strXml += '<a:ln><a:noFill/></a:ln>'
					} else {
						strXml += '<a:ln w="' + opts.lineSize * ONEPT + '" cap="flat"><a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>'
						strXml += '<a:prstDash val="' + (opts.lineDash || 'solid') + '"/><a:round/></a:ln>'
					}

					// Shadow
					strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW)
				}
				strXml += '  </c:spPr>'

				// 'c:marker' tag: `lineDataSymbol`
				{
					strXml += '<c:marker>'
					strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>'
					if (opts.lineDataSymbolSize) {
						// Defaults to "auto" otherwise (but this is usually too small, so there is a default)
						strXml += '  <c:size val="' + opts.lineDataSymbolSize + '"/>'
					}
					strXml += '  <c:spPr>'
					strXml +=
						'    <a:solidFill>' +
						createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]) +
						'</a:solidFill>'
					var symbolLineColor = opts.lineDataSymbolLineColor || strSerColor
					strXml +=
						'    <a:ln w="' +
						opts.lineDataSymbolLineSize +
						'" cap="flat"><a:solidFill>' +
						createColorElement(symbolLineColor) +
						'</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'
					strXml += '    <a:effectLst/>'
					strXml += '  </c:spPr>'
					strXml += '</c:marker>'
				}

				// Option: scatter data point labels
				if (opts.showLabel) {
					var chartUuid = getUuid('-xxxx-xxxx-xxxx-xxxxxxxxxxxx')
					if (obj.labels && (opts.dataLabelFormatScatter == 'custom' || opts.dataLabelFormatScatter == 'customXY')) {
						strXml += '<c:dLbls>'
						obj.labels.forEach((label, idx) => {
							if (opts.dataLabelFormatScatter == 'custom' || opts.dataLabelFormatScatter == 'customXY') {
								strXml += '  <c:dLbl>'
								strXml += '    <c:idx val="' + idx + '"/>'
								strXml += '    <c:tx>'
								strXml += '      <c:rich>'
								strXml += '			<a:bodyPr>'
								strXml += '				<a:spAutoFit/>'
								strXml += '			</a:bodyPr>'
								strXml += '        	<a:lstStyle/>'
								strXml += '        	<a:p>'
								strXml += '				<a:pPr>'
								strXml += '					<a:defRPr/>'
								strXml += '				</a:pPr>'
								strXml += '          	<a:r>'
								strXml += '            		<a:rPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>'
								strXml += '            		<a:t>' + encodeXmlEntities(label) + '</a:t>'
								strXml += '          	</a:r>'
								// Apply XY values at end of custom label
								// Do not apply the values if the label was empty or just spaces
								// This allows for selective labelling where required
								if (opts.dataLabelFormatScatter == 'customXY' && !/^ *$/.test(label)) {
									strXml += '          	<a:r>'
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>'
									strXml += '          		<a:t> (</a:t>'
									strXml += '          	</a:r>'
									strXml += '          	<a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="XVALUE">'
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>'
									strXml += '          		<a:pPr>'
									strXml += '          			<a:defRPr/>'
									strXml += '          		</a:pPr>'
									strXml += '          		<a:t>[' + encodeXmlEntities(obj.name) + '</a:t>'
									strXml += '          	</a:fld>'
									strXml += '          	<a:r>'
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>'
									strXml += '          		<a:t>, </a:t>'
									strXml += '          	</a:r>'
									strXml += '          	<a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="YVALUE">'
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>'
									strXml += '          		<a:pPr>'
									strXml += '          			<a:defRPr/>'
									strXml += '          		</a:pPr>'
									strXml += '          		<a:t>[' + encodeXmlEntities(obj.name) + ']</a:t>'
									strXml += '          	</a:fld>'
									strXml += '          	<a:r>'
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>'
									strXml += '          		<a:t>)</a:t>'
									strXml += '          	</a:r>'
									strXml += '          	<a:endParaRPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>'
								}
								strXml += '        	</a:p>'
								strXml += '      </c:rich>'
								strXml += '    </c:tx>'
								strXml += '    <c:spPr>'
								strXml += '    	<a:noFill/>'
								strXml += '    	<a:ln>'
								strXml += '    		<a:noFill/>'
								strXml += '    	</a:ln>'
								strXml += '    	<a:effectLst/>'
								strXml += '    </c:spPr>'
								strXml += '    <c:showLegendKey val="0"/>'
								strXml += '    <c:showVal val="0"/>'
								strXml += '    <c:showCatName val="0"/>'
								strXml += '    <c:showSerName val="0"/>'
								strXml += '    <c:showPercent val="0"/>'
								strXml += '    <c:showBubbleSize val="0"/>'
								strXml += '	  <c:showLeaderLines val="1"/>'
								strXml += '    <c:extLst>'
								strXml += '      <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">'
								strXml += '			<c15:dlblFieldTable/>'
								strXml += '			<c15:showDataLabelsRange val="0"/>'
								strXml += '		</c:ext>'
								strXml += '      <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
								strXml += '			<c16:uniqueId val="{' + '00000000'.substring(0, 8 - (idx + 1).toString().length).toString() + (idx + 1) + chartUuid + '}"/>'
								strXml += '      </c:ext>'
								strXml += '		</c:extLst>'
								strXml += '</c:dLbl>'
							}
						})
						strXml += '</c:dLbls>'
					}
					if (opts.dataLabelFormatScatter == 'XY') {
						strXml += '<c:dLbls>'
						strXml += '	<c:spPr>'
						strXml += '		<a:noFill/>'
						strXml += '		<a:ln>'
						strXml += '			<a:noFill/>'
						strXml += '		</a:ln>'
						strXml += '	  	<a:effectLst/>'
						strXml += '	</c:spPr>'
						strXml += '	<c:txPr>'
						strXml += '		<a:bodyPr>'
						strXml += '			<a:spAutoFit/>'
						strXml += '		</a:bodyPr>'
						strXml += '		<a:lstStyle/>'
						strXml += '		<a:p>'
						strXml += '	    	<a:pPr>'
						strXml += '        		<a:defRPr/>'
						strXml += '	    	</a:pPr>'
						strXml += '	    	<a:endParaRPr lang="en-US"/>'
						strXml += '		</a:p>'
						strXml += '	</c:txPr>'
						strXml += '	<c:showLegendKey val="0"/>'
						strXml += '	<c:showVal val="' + opts.showLabel ? '1' : '0' + '"/>'
						strXml += '	<c:showCatName val="' + opts.showLabel ? '1' : '0' + '"/>'
						strXml += '	<c:showSerName val="0"/>'
						strXml += '	<c:showPercent val="0"/>'
						strXml += '	<c:showBubbleSize val="0"/>'
						strXml += '	<c:extLst>'
						strXml += '		<c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">'
						strXml += '			<c15:showLeaderLines val="1"/>'
						strXml += '		</c:ext>'
						strXml += '	</c:extLst>'
						strXml += '</c:dLbls>'
					}
				}

				// Color bar chart bars various colors
				// Allow users with a single data set to pass their own array of colors (check for this using != ours)
				if ((data.length === 1 || opts.valueBarColors) && opts.chartColors != BARCHART_COLORS) {
					// Series Data Point colors
					obj.values.forEach((value, index) => {
						var arrColors = value < 0 ? opts.invertedColors || BARCHART_COLORS : opts.chartColors

						strXml += '  <c:dPt>'
						strXml += '    <c:idx val="' + index + '"/>'
						strXml += '      <c:invertIfNegative val="' + (opts.invertedColors ? 0 : 1) + '"/>'
						strXml += '    <c:bubble3D val="0"/>'
						strXml += '    <c:spPr>'
						if (opts.lineSize === 0) {
							strXml += '<a:ln><a:noFill/></a:ln>'
						} else {
							strXml += '<a:solidFill>'
							strXml += ' <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>'
							strXml += '</a:solidFill>'
						}
						strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW)
						strXml += '    </c:spPr>'
						strXml += '  </c:dPt>'
					})
				}

				// 3: "Values": Scatter Chart has 2: `xVal` and `yVal`
				{
					// X-Axis is always the same
					strXml += '<c:xVal>'
					strXml += '  <c:numRef>'
					strXml += '    <c:f>Sheet1!$A$2:$A$' + (data[0].values.length + 1) + '</c:f>'
					strXml += '    <c:numCache>'
					strXml += '      <c:formatCode>General</c:formatCode>'
					strXml += '      <c:ptCount val="' + data[0].values.length + '"/>'
					data[0].values.forEach((value, idx) => {
						strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value == 0 ? value : '') + '</c:v></c:pt>'
					})
					strXml += '    </c:numCache>'
					strXml += '  </c:numRef>'
					strXml += '</c:xVal>'

					// Y-Axis vals are this object's `values`
					strXml += '<c:yVal>'
					strXml += '  <c:numRef>'
					strXml += '    <c:f>Sheet1!$' + getExcelColName(idx + 1) + '$2:$' + getExcelColName(idx + 1) + '$' + (data[0].values.length + 1) + '</c:f>'
					strXml += '    <c:numCache>'
					strXml += '      <c:formatCode>General</c:formatCode>'
					// NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
					strXml += '      <c:ptCount val="' + data[0].values.length + '"/>'
					data[0].values.forEach((_value, idx) => {
						strXml += '<c:pt idx="' + idx + '"><c:v>' + (obj.values[idx] || obj.values[idx] == 0 ? obj.values[idx] : '') + '</c:v></c:pt>'
					})
					strXml += '    </c:numCache>'
					strXml += '  </c:numRef>'
					strXml += '</c:yVal>'
				}

				// Option: `smooth`
				strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>'

				// 4: Close "SERIES"
				strXml += '</c:ser>'
			})

			// 3: Data Labels
			{
				strXml += '  <c:dLbls>'
				strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>'
				strXml += '    <c:txPr>'
				strXml += '      <a:bodyPr/>'
				strXml += '      <a:lstStyle/>'
				strXml += '      <a:p><a:pPr>'
				strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">'
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>'
				strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>'
				strXml += '        </a:defRPr>'
				strXml += '      </a:pPr></a:p>'
				strXml += '    </c:txPr>'
				strXml += '    <c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>'
				strXml += '    <c:showLegendKey val="0"/>'
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
				strXml += '    <c:showCatName val="0"/>'
				strXml += '    <c:showSerName val="0"/>'
				strXml += '    <c:showPercent val="0"/>'
				strXml += '    <c:showBubbleSize val="0"/>'
				strXml += '  </c:dLbls>'
			}

			// 4: Add axisId (NOTE: order matters! (category comes first))
			strXml += '  <c:axId val="' + catAxisId + '"/>'
			strXml += '  <c:axId val="' + valAxisId + '"/>'

			// 5: Close Chart tag
			strXml += '</c:' + chartType + 'Chart>'

			// end switch
			break

		case 'bubble':
			/*
				`data` = [
					{ name:'X-Axis',     values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Values 1', values:[13, 20, 21, 25], sizes:[10, 5, 20, 15] },
					{ name:'Y-Values 2', values:[ 1,  2,  5,  9], sizes:[ 5, 3,  9,  3] }
				];
			*/

			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>'
			strXml += '<c:varyColors val="0"/>'

			// 2: Series: (One for each Y-Axis)
			var colorIndex = -1
			var idxColLtr = 1
			data.filter((_obj, idx) => {
				return idx > 0
			}).forEach((obj, idx) => {
				colorIndex++
				strXml += '<c:ser>'
				strXml += '  <c:idx val="' + idx + '"/>'
				strXml += '  <c:order val="' + idx + '"/>'

				// A: `<c:tx>`
				strXml += '  <c:tx>'
				strXml += '    <c:strRef>'
				strXml += '      <c:f>Sheet1!$' + LETTERS[idxColLtr] + '$1</c:f>'
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + obj.name + '</c:v></c:pt></c:strCache>'
				strXml += '    </c:strRef>'
				strXml += '  </c:tx>'

				// B: '<c:spPr>': Fill, Border, Line, LineStyle (dash, etc.), Shadow
				{
					strXml += '<c:spPr>'

					var strSerColor = opts.chartColors[colorIndex % opts.chartColors.length]

					if (strSerColor == 'transparent') {
						strXml += '<a:noFill/>'
					} else if (opts.chartColorsOpacity) {
						strXml += '<a:solidFill>' + createColorElement(strSerColor, '<a:alpha val="' + opts.chartColorsOpacity + '000"/>') + '</a:solidFill>'
					} else {
						strXml += '<a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>'
					}

					if (opts.lineSize == 0) {
						strXml += '<a:ln><a:noFill/></a:ln>'
					} else if (opts.dataBorder) {
						strXml +=
							'<a:ln w="' +
							opts.dataBorder.pt * ONEPT +
							'" cap="flat"><a:solidFill>' +
							createColorElement(opts.dataBorder.color) +
							'</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'
					} else {
						strXml += '<a:ln w="' + opts.lineSize * ONEPT + '" cap="flat"><a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>'
						strXml += '<a:prstDash val="' + (opts.lineDash || 'solid') + '"/><a:round/></a:ln>'
					}

					// Shadow
					strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW)

					strXml += '</c:spPr>'
				}

				// C: '<c:dLbls>' "Data Labels"
				// Let it be defaulted for now

				// D: '<c:xVal>'/'<c:yVal>' "Values": Scatter Chart has 2: `xVal` and `yVal`
				{
					// X-Axis is always the same
					strXml += '<c:xVal>'
					strXml += '  <c:numRef>'
					strXml += '    <c:f>Sheet1!$A$2:$A$' + (data[0].values.length + 1) + '</c:f>'
					strXml += '    <c:numCache>'
					strXml += '      <c:formatCode>General</c:formatCode>'
					strXml += '      <c:ptCount val="' + data[0].values.length + '"/>'
					data[0].values.forEach((value, idx) => {
						strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value == 0 ? value : '') + '</c:v></c:pt>'
					})
					strXml += '    </c:numCache>'
					strXml += '  </c:numRef>'
					strXml += '</c:xVal>'

					// Y-Axis vals are this object's `values`
					strXml += '<c:yVal>'
					strXml += '  <c:numRef>'
					strXml += '    <c:f>Sheet1!$' + getExcelColName(idxColLtr) + '$2:$' + getExcelColName(idxColLtr) + '$' + (data[0].values.length + 1) + '</c:f>'
					idxColLtr++
					strXml += '    <c:numCache>'
					strXml += '      <c:formatCode>General</c:formatCode>'
					// NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
					strXml += '      <c:ptCount val="' + data[0].values.length + '"/>'
					data[0].values.forEach((_value, idx) => {
						strXml += '<c:pt idx="' + idx + '"><c:v>' + (obj.values[idx] || obj.values[idx] == 0 ? obj.values[idx] : '') + '</c:v></c:pt>'
					})
					strXml += '    </c:numCache>'
					strXml += '  </c:numRef>'
					strXml += '</c:yVal>'
				}

				// E: '<c:bubbleSize>'
				strXml += '  <c:bubbleSize>'
				strXml += '    <c:numRef>'
				strXml += '      <c:f>Sheet1!' + '$' + getExcelColName(idxColLtr) + '$2:$' + getExcelColName(idx + 2) + '$' + (obj.sizes.length + 1) + '</c:f>'
				idxColLtr++
				strXml += '      <c:numCache>'
				strXml += '        <c:formatCode>General</c:formatCode>'
				strXml += '	       <c:ptCount val="' + obj.sizes.length + '"/>'
				obj.sizes.forEach((value, idx) => {
					strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || '') + '</c:v></c:pt>'
				})
				strXml += '      </c:numCache>'
				strXml += '    </c:numRef>'
				strXml += '  </c:bubbleSize>'
				strXml += '  <c:bubble3D val="0"/>'

				// F: Close "SERIES"
				strXml += '</c:ser>'
			})

			// 3: Data Labels
			{
				strXml += '  <c:dLbls>'
				strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>'
				strXml += '    <c:txPr>'
				strXml += '      <a:bodyPr/>'
				strXml += '      <a:lstStyle/>'
				strXml += '      <a:p><a:pPr>'
				strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">'
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>'
				strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>'
				strXml += '        </a:defRPr>'
				strXml += '      </a:pPr></a:p>'
				strXml += '    </c:txPr>'
				strXml += '    <c:dLblPos val="ctr"/>'
				strXml += '    <c:showLegendKey val="0"/>'
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
				strXml += '    <c:showCatName val="0"/>'
				strXml += '    <c:showSerName val="0"/>'
				strXml += '    <c:showPercent val="0"/>'
				strXml += '    <c:showBubbleSize val="0"/>'
				strXml += '  </c:dLbls>'
			}

			// 4: Add bubble options
			//strXml += '  <c:bubbleScale val="100"/>';
			//strXml += '  <c:showNegBubbles val="0"/>';
			// Commented out to let it default to PPT until we create options

			// 5: Add axisId (NOTE: order matters! (category comes first))
			strXml += '  <c:axId val="' + catAxisId + '"/>'
			strXml += '  <c:axId val="' + valAxisId + '"/>'

			// 6: Close Chart tag
			strXml += '</c:' + chartType + 'Chart>'

			// end switch
			break

		case 'pie':
		case 'doughnut':
			// Use the same var name so code blocks from barChart are interchangeable
			var obj = data[0]

			/* EX:
				data: [
				 {
				   name: 'Project Status',
				   labels: ['Red', 'Amber', 'Green', 'Unknown'],
				   values: [10, 20, 38, 2]
				 }
				]
			*/

			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>'
			strXml += '  <c:varyColors val="0"/>'
			strXml += '<c:ser>'
			strXml += '  <c:idx val="0"/>'
			strXml += '  <c:order val="0"/>'
			strXml += '  <c:tx>'
			strXml += '    <c:strRef>'
			strXml += '      <c:f>Sheet1!$B$1</c:f>'
			strXml += '      <c:strCache>'
			strXml += '        <c:ptCount val="1"/>'
			strXml += '        <c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt>'
			strXml += '      </c:strCache>'
			strXml += '    </c:strRef>'
			strXml += '  </c:tx>'
			strXml += '  <c:spPr>'
			strXml += '    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'
			strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'
			if (opts.dataNoEffects) {
				strXml += '<a:effectLst/>'
			} else {
				strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW)
			}
			strXml += '  </c:spPr>'
			strXml += '<c:explosion val="0"/>'

			// 2: "Data Point" block for every data row
			obj.labels.forEach((_label, idx) => {
				strXml += '<c:dPt>'
				strXml += '  <c:idx val="' + idx + '"/>'
				strXml += '  <c:explosion val="0"/>'
				strXml += '  <c:spPr>'
				strXml +=
					'    <a:solidFill>' +
					createColorElement(opts.chartColors[idx + 1 > opts.chartColors.length ? Math.floor(Math.random() * opts.chartColors.length) : idx]) +
					'</a:solidFill>'
				if (opts.dataBorder) {
					strXml +=
						'<a:ln w="' +
						opts.dataBorder.pt * ONEPT +
						'" cap="flat"><a:solidFill>' +
						createColorElement(opts.dataBorder.color) +
						'</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'
				}
				strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW)
				strXml += '  </c:spPr>'
				strXml += '</c:dPt>'
			})

			// 3: "Data Label" block for every data Label
			strXml += '<c:dLbls>'
			obj.labels.forEach((_label, idx) => {
				strXml += '<c:dLbl>'
				strXml += '  <c:idx val="' + idx + '"/>'
				strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>'
				strXml += '    <c:txPr>'
				strXml += '      <a:bodyPr/><a:lstStyle/>'
				strXml += '      <a:p><a:pPr>'
				strXml +=
					'        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">'
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>'
				strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>'
				strXml += '        </a:defRPr>'
				strXml += '      </a:pPr></a:p>'
				strXml += '    </c:txPr>'
				if (chartType == 'pie') {
					strXml += '    <c:dLblPos val="' + (opts.dataLabelPosition || 'inEnd') + '"/>'
				}
				strXml += '    <c:showLegendKey val="0"/>'
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
				strXml += '    <c:showCatName val="' + (opts.showLabel ? '1' : '0') + '"/>'
				strXml += '    <c:showSerName val="0"/>'
				strXml += '    <c:showPercent val="' + (opts.showPercent ? '1' : '0') + '"/>'
				strXml += '    <c:showBubbleSize val="0"/>'
				strXml += '  </c:dLbl>'
			})
			strXml +=
				'<c:numFmt formatCode="' +
				opts.dataLabelFormatCode +
				'" sourceLinked="0"/>\
				<c:txPr>\
				  <a:bodyPr/>\
				  <a:lstStyle/>\
				  <a:p>\
					<a:pPr>\
					  <a:defRPr b="0" i="0" strike="noStrike" sz="1800" u="none">\
						<a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/>\
					  </a:defRPr>\
					</a:pPr>\
				  </a:p>\
				</c:txPr>\
				' +
				(chartType == 'pie' ? '<c:dLblPos val="ctr"/>' : '') +
				'\
				<c:showLegendKey val="0"/>\
				<c:showVal val="0"/>\
				<c:showCatName val="1"/>\
				<c:showSerName val="0"/>\
				<c:showPercent val="1"/>\
				<c:showBubbleSize val="0"/>\
				<c:showLeaderLines val="0"/>'
			strXml += '</c:dLbls>'

			// 2: "Categories"
			strXml += '<c:cat>'
			strXml += '  <c:strRef>'
			strXml += '    <c:f>Sheet1!' + '$A$2:$A$' + (obj.labels.length + 1) + '</c:f>'
			strXml += '    <c:strCache>'
			strXml += '	     <c:ptCount val="' + obj.labels.length + '"/>'
			obj.labels.forEach((label, idx) => {
				strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>'
			})
			strXml += '    </c:strCache>'
			strXml += '  </c:strRef>'
			strXml += '</c:cat>'

			// 3: Create vals
			strXml += '  <c:val>'
			strXml += '    <c:numRef>'
			strXml += '      <c:f>Sheet1!' + '$B$2:$B$' + (obj.labels.length + 1) + '</c:f>'
			strXml += '      <c:numCache>'
			strXml += '	       <c:ptCount val="' + obj.labels.length + '"/>'
			obj.values.forEach((value, idx) => {
				strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value == 0 ? value : '') + '</c:v></c:pt>'
			})
			strXml += '      </c:numCache>'
			strXml += '    </c:numRef>'
			strXml += '  </c:val>'

			// 4: Close "SERIES"
			strXml += '  </c:ser>'
			strXml += '  <c:firstSliceAng val="0"/>'
			if (chartType == 'doughnut') strXml += '  <c:holeSize val="' + (opts.holeSize || 50) + '"/>'
			strXml += '</c:' + chartType + 'Chart>'

			// Done with Doughnut/Pie
			break
	}

	return strXml
}

function makeCatAxis(opts, axisId, valAxisId) {
	var strXml = ''

	// Build cat axis tag
	// NOTE: Scatter and Bubble chart need two Val axises as they display numbers on x axis
	if (opts.type == 'scatter' || opts.type == 'bubble') {
		strXml += '<c:valAx>'
	} else {
		strXml += '<c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>'
	}
	strXml += '  <c:axId val="' + axisId + '"/>'
	strXml += '  <c:scaling>'
	strXml += '<c:orientation val="' + (opts.catAxisOrientation || (opts.barDir == 'col' ? 'minMax' : 'minMax')) + '"/>'
	if (opts.catAxisMaxVal || opts.catAxisMaxVal == 0) strXml += '<c:max val="' + opts.catAxisMaxVal + '"/>'
	if (opts.catAxisMinVal || opts.catAxisMinVal == 0) strXml += '<c:min val="' + opts.catAxisMinVal + '"/>'
	strXml += '</c:scaling>'
	strXml += '  <c:delete val="' + (opts.catAxisHidden ? 1 : 0) + '"/>'
	strXml += '  <c:axPos val="' + (opts.barDir == 'col' ? 'b' : 'l') + '"/>'
	strXml += opts.catGridLine !== 'none' ? createGridLineElement(opts.catGridLine) : ''
	// '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
	if (opts.showCatAxisTitle) {
		strXml += genXmlTitle({
			color: opts.catAxisTitleColor,
			fontFace: opts.catAxisTitleFontFace,
			fontSize: opts.catAxisTitleFontSize,
			rotate: opts.catAxisTitleRotate,
			title: opts.catAxisTitle || 'Axis Title',
		})
	}
	// NOTE: Adding Val Axis Formatting if scatter or bubble charts
	if (opts.type == 'scatter' || opts.type == 'bubble') {
		strXml += '  <c:numFmt formatCode="' + (opts.valAxisLabelFormatCode ? opts.valAxisLabelFormatCode : 'General') + '" sourceLinked="0"/>'
	} else {
		strXml += '  <c:numFmt formatCode="' + (opts.catLabelFormatCode || 'General') + '" sourceLinked="0"/>'
	}
	if (opts.type === 'scatter') {
		strXml += '  <c:majorTickMark val="none"/>'
		strXml += '  <c:minorTickMark val="none"/>'
		strXml += '  <c:tickLblPos val="nextTo"/>'
	} else {
		strXml += '  <c:majorTickMark val="out"/>'
		strXml += '  <c:minorTickMark val="none"/>'
		strXml += '  <c:tickLblPos val="' + (opts.catAxisLabelPos || opts.barDir == 'col' ? 'low' : 'nextTo') + '"/>'
	}
	strXml += '  <c:spPr>'
	strXml += '    <a:ln w="12700" cap="flat">'
	strXml += opts.catAxisLineShow == false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + DEF_CHART_GRIDLINE.color + '"/></a:solidFill>'
	strXml += '      <a:prstDash val="solid"/>'
	strXml += '      <a:round/>'
	strXml += '    </a:ln>'
	strXml += '  </c:spPr>'
	strXml += '  <c:txPr>'
	strXml += '    <a:bodyPr ' + (opts.catAxisLabelRotate ? 'rot="' + convertRotationDegrees(opts.catAxisLabelRotate) + '"' : '') + '/>' // don't specify rot 0 so we get the auto behavior
	strXml += '    <a:lstStyle/>'
	strXml += '    <a:p>'
	strXml += '    <a:pPr>'
	strXml += '    <a:defRPr sz="' + (opts.catAxisLabelFontSize || DEF_FONT_SIZE) + '00" b="' + (opts.catAxisLabelFontBold ? 1 : 0) + '" i="0" u="none" strike="noStrike">'
	strXml += '      <a:solidFill><a:srgbClr val="' + (opts.catAxisLabelColor || DEF_FONT_COLOR) + '"/></a:solidFill>'
	strXml += '      <a:latin typeface="' + (opts.catAxisLabelFontFace || 'Arial') + '"/>'
	strXml += '   </a:defRPr>'
	strXml += '  </a:pPr>'
	strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>'
	strXml += '  </a:p>'
	strXml += ' </c:txPr>'
	strXml += ' <c:crossAx val="' + valAxisId + '"/>'
	strXml += ' <c:' + (typeof opts.valAxisCrossesAt === 'number' ? 'crossesAt' : 'crosses') + ' val="' + opts.valAxisCrossesAt + '"/>'
	strXml += ' <c:auto val="1"/>'
	strXml += ' <c:lblAlgn val="ctr"/>'
	strXml += ' <c:noMultiLvlLbl val="1"/>'
	if (opts.catAxisLabelFrequency) strXml += ' <c:tickLblSkip val="' + opts.catAxisLabelFrequency + '"/>'

	// Issue#149: PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
	if (opts.catLabelFormatCode) {
		;['catAxisBaseTimeUnit', 'catAxisMajorTimeUnit', 'catAxisMinorTimeUnit'].forEach((opt, idx) => {
			// Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
			if (opts[opt] && (typeof opts[opt] !== 'string' || ['days', 'months', 'years'].indexOf(opt.toLowerCase()) == -1)) {
				console.warn('`' + opt + "` must be one of: 'days','months','years' !")
				opts[opt] = null
			}
		})
		if (opts.catAxisBaseTimeUnit) strXml += ' <c:baseTimeUnit  val="' + opts.catAxisBaseTimeUnit.toLowerCase() + '"/>'
		if (opts.catAxisMajorTimeUnit) strXml += ' <c:majorTimeUnit val="' + opts.catAxisMajorTimeUnit.toLowerCase() + '"/>'
		if (opts.catAxisMinorTimeUnit) strXml += ' <c:minorTimeUnit val="' + opts.catAxisMinorTimeUnit.toLowerCase() + '"/>'
		if (opts.catAxisMajorUnit) strXml += ' <c:majorUnit     val="' + opts.catAxisMajorUnit + '"/>'
		if (opts.catAxisMinorUnit) strXml += ' <c:minorUnit     val="' + opts.catAxisMinorUnit + '"/>'
	}

	// Close cat axis tag
	// NOTE: Added closing tag of val or cat axis based on chart type
	if (opts.type == 'scatter' || opts.type == 'bubble') {
		strXml += '</c:valAx>'
	} else {
		strXml += '</c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>'
	}

	return strXml
}

function makeValueAxis(opts: ISlideRelChart['opts'], valAxisId: string) {
	var axisPos = valAxisId === AXIS_ID_VALUE_PRIMARY ? (opts.barDir == 'col' ? 'l' : 'b') : opts.barDir == 'col' ? 'r' : 't'
	var strXml = ''
	var isRight = axisPos === 'r' || axisPos === 't'
	var crosses = isRight ? 'max' : 'autoZero'
	var crossAxId = valAxisId === AXIS_ID_VALUE_PRIMARY ? AXIS_ID_CATEGORY_PRIMARY : AXIS_ID_CATEGORY_SECONDARY

	strXml += '<c:valAx>'
	strXml += '  <c:axId val="' + valAxisId + '"/>'
	strXml += '  <c:scaling>'
	strXml += '    <c:orientation val="' + (opts.valAxisOrientation || (opts.barDir == 'col' ? 'minMax' : 'minMax')) + '"/>'
	if (opts.valAxisMaxVal || opts.valAxisMaxVal == 0) strXml += '<c:max val="' + opts.valAxisMaxVal + '"/>'
	if (opts.valAxisMinVal || opts.valAxisMinVal == 0) strXml += '<c:min val="' + opts.valAxisMinVal + '"/>'
	strXml += '  </c:scaling>'
	strXml += '  <c:delete val="' + (opts.valAxisHidden ? 1 : 0) + '"/>'
	strXml += '  <c:axPos val="' + axisPos + '"/>'
	if (opts.valGridLine != 'none') strXml += createGridLineElement(opts.valGridLine)
	// '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
	if (opts.showValAxisTitle) {
		strXml += genXmlTitle({
			color: opts.valAxisTitleColor,
			fontFace: opts.valAxisTitleFontFace,
			fontSize: opts.valAxisTitleFontSize,
			rotate: opts.valAxisTitleRotate,
			title: opts.valAxisTitle || 'Axis Title',
		})
	}
	strXml += ' <c:numFmt formatCode="' + (opts.valAxisLabelFormatCode ? opts.valAxisLabelFormatCode : 'General') + '" sourceLinked="0"/>'
	if (opts.type === 'scatter') {
		strXml += '  <c:majorTickMark val="none"/>'
		strXml += '  <c:minorTickMark val="none"/>'
		strXml += '  <c:tickLblPos val="nextTo"/>'
	} else {
		strXml += ' <c:majorTickMark val="out"/>'
		strXml += ' <c:minorTickMark val="none"/>'
		strXml += ' <c:tickLblPos val="' + (opts.catAxisLabelPos || opts.barDir == 'col' ? 'nextTo' : 'low') + '"/>'
	}
	strXml += ' <c:spPr>'
	strXml += '   <a:ln w="12700" cap="flat">'
	strXml += opts.valAxisLineShow == false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + DEF_CHART_GRIDLINE.color + '"/></a:solidFill>'
	strXml += '     <a:prstDash val="solid"/>'
	strXml += '     <a:round/>'
	strXml += '   </a:ln>'
	strXml += ' </c:spPr>'
	strXml += ' <c:txPr>'
	strXml += '  <a:bodyPr ' + (opts.valAxisLabelRotate ? 'rot="' + convertRotationDegrees(opts.valAxisLabelRotate) + '"' : '') + '/>' // don't specify rot 0 so we get the auto behavior
	strXml += '  <a:lstStyle/>'
	strXml += '  <a:p>'
	strXml += '    <a:pPr>'
	strXml += '      <a:defRPr sz="' + (opts.valAxisLabelFontSize || DEF_FONT_SIZE) + '00" b="' + (opts.valAxisLabelFontBold ? 1 : 0) + '" i="0" u="none" strike="noStrike">'
	strXml += '        <a:solidFill><a:srgbClr val="' + (opts.valAxisLabelColor || DEF_FONT_COLOR) + '"/></a:solidFill>'
	strXml += '        <a:latin typeface="' + (opts.valAxisLabelFontFace || 'Arial') + '"/>'
	strXml += '      </a:defRPr>'
	strXml += '    </a:pPr>'
	strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>'
	strXml += '  </a:p>'
	strXml += ' </c:txPr>'
	strXml += ' <c:crossAx val="' + crossAxId + '"/>'
	strXml += ' <c:crosses val="' + crosses + '"/>'
	strXml +=
		' <c:crossBetween val="' +
		(opts.type === 'scatter' || (Array.isArray(opts.type) && opts.type.indexOf(CHART_TYPES.AREA) > -1 ? true : false) ? 'midCat' : 'between') +
		'"/>'
	if (opts.valAxisMajorUnit) strXml += ' <c:majorUnit val="' + opts.valAxisMajorUnit + '"/>'
	strXml += '</c:valAx>'

	return strXml
}

/**
 * @description Used by `bar3D`
 * @param `opts` (ISlideRelChart['opts']) - chart options
 * @param `axisId` (string)
 * @param `valAxisId` (string)
 */
function makeSerAxis(opts: ISlideRelChart['opts'], axisId: string, valAxisId: string) {
	var strXml = ''

	// Build ser axis tag
	strXml += '<c:serAx>'
	strXml += '  <c:axId val="' + axisId + '"/>'
	strXml += '  <c:scaling><c:orientation val="' + (opts.serAxisOrientation || (opts.barDir == 'col' ? 'minMax' : 'minMax')) + '"/></c:scaling>'
	strXml += '  <c:delete val="' + (opts.serAxisHidden ? 1 : 0) + '"/>'
	strXml += '  <c:axPos val="' + (opts.barDir == 'col' ? 'b' : 'l') + '"/>'
	strXml += opts.serGridLine !== 'none' ? createGridLineElement(opts.serGridLine) : ''
	// '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
	if (opts.showSerAxisTitle) {
		strXml += genXmlTitle({
			color: opts.serAxisTitleColor,
			fontFace: opts.serAxisTitleFontFace,
			fontSize: opts.serAxisTitleFontSize,
			rotate: opts.serAxisTitleRotate,
			title: opts.serAxisTitle || 'Axis Title',
		})
	}
	strXml += '  <c:numFmt formatCode="' + (opts.serLabelFormatCode || 'General') + '" sourceLinked="0"/>'
	strXml += '  <c:majorTickMark val="out"/>'
	strXml += '  <c:minorTickMark val="none"/>'
	strXml += '  <c:tickLblPos val="' + (opts.serAxisLabelPos || opts.barDir == 'col' ? 'low' : 'nextTo') + '"/>'
	strXml += '  <c:spPr>'
	strXml += '    <a:ln w="12700" cap="flat">'
	strXml += opts.serAxisLineShow == false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + DEF_CHART_GRIDLINE.color + '"/></a:solidFill>'
	strXml += '      <a:prstDash val="solid"/>'
	strXml += '      <a:round/>'
	strXml += '    </a:ln>'
	strXml += '  </c:spPr>'
	strXml += '  <c:txPr>'
	strXml += '    <a:bodyPr/>' // don't specify rot 0 so we get the auto behavior
	strXml += '    <a:lstStyle/>'
	strXml += '    <a:p>'
	strXml += '    <a:pPr>'
	strXml += '    <a:defRPr sz="' + (opts.serAxisLabelFontSize || DEF_FONT_SIZE) + '00" b="0" i="0" u="none" strike="noStrike">'
	strXml += '      <a:solidFill><a:srgbClr val="' + (opts.serAxisLabelColor || DEF_FONT_COLOR) + '"/></a:solidFill>'
	strXml += '      <a:latin typeface="' + (opts.serAxisLabelFontFace || 'Arial') + '"/>'
	strXml += '   </a:defRPr>'
	strXml += '  </a:pPr>'
	strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>'
	strXml += '  </a:p>'
	strXml += ' </c:txPr>'
	strXml += ' <c:crossAx val="' + valAxisId + '"/>'
	strXml += ' <c:crosses val="autoZero"/>'
	if (opts.serAxisLabelFrequency) strXml += ' <c:tickLblSkip val="' + opts.serAxisLabelFrequency + '"/>'

	// Issue#149: PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
	if (opts.serLabelFormatCode) {
		;['serAxisBaseTimeUnit', 'serAxisMajorTimeUnit', 'serAxisMinorTimeUnit'].forEach(opt => {
			// Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
			if (opts[opt] && (typeof opts[opt] !== 'string' || ['days', 'months', 'years'].indexOf(opt.toLowerCase()) == -1)) {
				console.warn('`' + opt + "` must be one of: 'days','months','years' !")
				opts[opt] = null
			}
		})
		if (opts.serAxisBaseTimeUnit) strXml += ' <c:baseTimeUnit  val="' + opts.serAxisBaseTimeUnit.toLowerCase() + '"/>'
		if (opts.serAxisMajorTimeUnit) strXml += ' <c:majorTimeUnit val="' + opts.serAxisMajorTimeUnit.toLowerCase() + '"/>'
		if (opts.serAxisMinorTimeUnit) strXml += ' <c:minorTimeUnit val="' + opts.serAxisMinorTimeUnit.toLowerCase() + '"/>'
		if (opts.serAxisMajorUnit) strXml += ' <c:majorUnit     val="' + opts.serAxisMajorUnit + '"/>'
		if (opts.serAxisMinorUnit) strXml += ' <c:minorUnit     val="' + opts.serAxisMinorUnit + '"/>'
	}

	// Close ser axis tag
	strXml += '</c:serAx>'

	return strXml
}

/**
 * DESC: Generate the XML for title elements used for the char and axis titles
 */
function genXmlTitle(opts) {
	var align = opts.titleAlign == 'left' ? 'l' : opts.titleAlign == 'right' ? 'r' : false
	var strXml = ''
	strXml += '<c:title>'
	strXml += ' <c:tx>'
	strXml += '  <c:rich>'
	if (opts.rotate) {
		strXml += '  <a:bodyPr rot="' + convertRotationDegrees(opts.rotate) + '"/>'
	} else {
		strXml += '  <a:bodyPr/>' // don't specify rotation to get default (ex. vertical for cat axis)
	}
	strXml += '  <a:lstStyle/>'
	strXml += '  <a:p>'
	strXml += align ? '<a:pPr algn="' + align + '">' : '<a:pPr>'
	var sizeAttr = ''
	if (opts.fontSize) {
		// only set the font size if specified.  Powerpoint will handle the default size
		sizeAttr = 'sz="' + Math.round(opts.fontSize) + '00"'
	}
	strXml += '      <a:defRPr ' + sizeAttr + ' b="0" i="0" u="none" strike="noStrike">'
	strXml += '        <a:solidFill><a:srgbClr val="' + (opts.color || DEF_FONT_COLOR) + '"/></a:solidFill>'
	strXml += '        <a:latin typeface="' + (opts.fontFace || 'Arial') + '"/>'
	strXml += '      </a:defRPr>'
	strXml += '    </a:pPr>'
	strXml += '    <a:r>'
	strXml += '      <a:rPr ' + sizeAttr + ' b="0" i="0" u="none" strike="noStrike">'
	strXml += '        <a:solidFill><a:srgbClr val="' + (opts.color || DEF_FONT_COLOR) + '"/></a:solidFill>'
	strXml += '        <a:latin typeface="' + (opts.fontFace || 'Arial') + '"/>'
	strXml += '      </a:rPr>'
	strXml += '      <a:t>' + (encodeXmlEntities(opts.title) || '') + '</a:t>'
	strXml += '    </a:r>'
	strXml += '  </a:p>'
	strXml += '  </c:rich>'
	strXml += ' </c:tx>'
	if (opts.titlePos && opts.titlePos.x && opts.titlePos.y) {
		strXml += '<c:layout>'
		strXml += '  <c:manualLayout>'
		strXml += '    <c:xMode val="edge"/>'
		strXml += '    <c:yMode val="edge"/>'
		strXml += '    <c:x val="' + opts.titlePos.x + '"/>'
		strXml += '    <c:y val="' + opts.titlePos.y + '"/>'
		strXml += '  </c:manualLayout>'
		strXml += '</c:layout>'
	} else {
		strXml += ' <c:layout/>'
	}
	strXml += ' <c:overlay val="0"/>'
	strXml += '</c:title>'
	return strXml
}

/**
* DESC: Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
* EX:
	<p:txBody>
		<a:bodyPr wrap="none" lIns="50800" tIns="50800" rIns="50800" bIns="50800" anchor="ctr">
		</a:bodyPr>
		<a:lstStyle/>
		<a:p>
		  <a:pPr marL="228600" indent="-228600"><a:buSzPct val="100000"/><a:buChar char="&#x2022;"/></a:pPr>
		  <a:r>
			<a:t>bullet 1 </a:t>
		  </a:r>
		  <a:r>
			<a:rPr>
			  <a:solidFill><a:srgbClr val="7B2CD6"/></a:solidFill>
			</a:rPr>
			<a:t>colored text</a:t>
		  </a:r>
		</a:p>
	  </p:txBody>
* NOTES:
* - PPT text lines [lines followed by line-breaks] are createing using <p>-aragraph's
* - Bullets are a paragprah-level formatting device
*
* @param slideObj (object) - slideObj -OR- table `cell` object
* @returns XML string containing the param object's text and formatting
*/
export function genXmlTextBody(slideObj) {
	// FIRST: Shapes without text, etc. may be sent here during build, but have no text to render so return an empty string
	if (!slideObj.options.isTableCell && (typeof slideObj.text === 'undefined' || slideObj.text == null)) return ''

	// Create options if needed
	if (!slideObj.options) slideObj.options = {}

	// Vars
	var arrTextObjects = []
	var tagStart = slideObj.options.isTableCell ? '<a:txBody>' : '<p:txBody>'
	var tagClose = slideObj.options.isTableCell ? '</a:txBody>' : '</p:txBody>'
	var strSlideXml = tagStart

	// STEP 1: Modify slideObj to be consistent array of `{ text:'', options:{} }`
	/* CASES:
		addText( 'string' )
		addText( 'line1\n line2' )
		addText( ['barry','allen'] )
		addText( [{text'word1'}, {text:'word2'}] )
		addText( [{text'line1\n line2'}, {text:'end word'}] )
	*/
	// A: Handle string/number
	if (typeof slideObj.text === 'string' || typeof slideObj.text === 'number') {
		slideObj.text = [{ text: slideObj.text.toString(), options: slideObj.options || {} }]
	}

	// STEP 2: Grab options, format line-breaks, etc.
	if (Array.isArray(slideObj.text)) {
		slideObj.text.forEach((obj, idx) => {
			// A: Set options
			obj.options = obj.options || slideObj.options || {}
			if (idx == 0 && obj.options && !obj.options.bullet && slideObj.options.bullet) obj.options.bullet = slideObj.options.bullet

			// B: Cast to text-object and fix line-breaks (if needed)
			if (typeof obj.text === 'string' || typeof obj.text === 'number') {
				obj.text = obj.text.toString().replace(/\r*\n/g, CRLF)
				// Plain strings like "hello \n world" need to have lineBreaks set to break as intended
				if (obj.text.indexOf(CRLF) > -1) obj.options.breakLine = true
			}

			// C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
			if (obj.text.split(CRLF).length > 0) {
				obj.text
					.toString()
					.split(CRLF)
					.forEach((line, idx) => {
						// Add line-breaks if not bullets/aligned (we add CRLF for those below in STEP 2)
						line += obj.options.breakLine && !obj.options.bullet && !obj.options.align ? CRLF : ''
						arrTextObjects.push({ text: line, options: obj.options })
					})
			} else {
				// NOTE: The replace used here is for non-textObjects (plain strings) eg:'hello\nworld'
				arrTextObjects.push(obj)
			}
		})
	}

	// STEP 3: Add bodyProperties
	{
		// A: 'bodyPr'
		strSlideXml += genXmlBodyProperties(slideObj.options)

		// B: 'lstStyle'
		// NOTE: Shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
		// FIXME: LINE horiz-align doesnt work (text is always to the left inside line) (FYI: the PPT code diff is substantial!)
		if (slideObj.options.h == 0 && slideObj.options.line && slideObj.options.align) {
			strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>'
		} else if (slideObj.type === 'placeholder') {
			strSlideXml += '<a:lstStyle>'
			strSlideXml += genXmlParagraphProperties(slideObj, true)
			strSlideXml += '</a:lstStyle>'
		} else {
			strSlideXml += '<a:lstStyle/>'
		}
	}

	// STEP 4: Loop over each text object and create paragraph props, text run, etc.
	arrTextObjects.forEach((textObj, idx) => {
		// Clear/Increment loop vars
		paragraphPropXml = '<a:pPr ' + (textObj.options.rtlMode ? ' rtl="1" ' : '')
		textObj.options.lineIdx = idx

		// Inherit pPr-type options from parent shape's `options`
		textObj.options.align = textObj.options.align || slideObj.options.align
		textObj.options.lineSpacing = textObj.options.lineSpacing || slideObj.options.lineSpacing
		textObj.options.indentLevel = textObj.options.indentLevel || slideObj.options.indentLevel
		textObj.options.paraSpaceBefore = textObj.options.paraSpaceBefore || slideObj.options.paraSpaceBefore
		textObj.options.paraSpaceAfter = textObj.options.paraSpaceAfter || slideObj.options.paraSpaceAfter

		textObj.options.lineIdx = idx
		var paragraphPropXml = genXmlParagraphProperties(textObj, false)

		// B: Start paragraph if this is the first text obj, or if current textObj is about to be bulleted or aligned
		if (idx == 0) {
			// Add paragraphProperties right after <p> before textrun(s) begin
			strSlideXml += '<a:p>' + paragraphPropXml
		} else if (idx > 0 && (typeof textObj.options.bullet !== 'undefined' || typeof textObj.options.align !== 'undefined')) {
			strSlideXml += '</a:p><a:p>' + paragraphPropXml
		}

		// C: Inherit any main options (color, fontSize, etc.)
		// We only pass the text.options to genXmlTextRun (not the Slide.options),
		// so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
		// TODO-3: convert to Object.values or whatever in ES6
		jQuery.each(slideObj.options, (key, val) => {
			// NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
			if (key != 'bullet' && !textObj.options[key]) textObj.options[key] = val
		})

		// D: Add formatted textrun
		strSlideXml += genXmlTextRun(textObj.options, textObj.text)
	})

	// STEP 5: Append 'endParaRPr' (when needed) and close current open paragraph
	// NOTE: (ISSUE#20/#193): Add 'endParaRPr' with font/size props or PPT default (Arial/18pt en-us) is used making row "too tall"/not honoring opts
	if (slideObj.options.isTableCell && (slideObj.options.fontSize || slideObj.options.fontFace)) {
		strSlideXml +=
			'<a:endParaRPr lang="' +
			(slideObj.options.lang ? slideObj.options.lang : 'en-US') +
			'" ' +
			(slideObj.options.fontSize ? ' sz="' + Math.round(slideObj.options.fontSize) + '00"' : '') +
			' dirty="0">'
		if (slideObj.options.fontFace) {
			strSlideXml += '  <a:latin typeface="' + slideObj.options.fontFace + '" charset="0" />'
			strSlideXml += '  <a:ea    typeface="' + slideObj.options.fontFace + '" charset="0" />'
			strSlideXml += '  <a:cs    typeface="' + slideObj.options.fontFace + '" charset="0" />'
		}
		strSlideXml += '</a:endParaRPr>'
	} else {
		strSlideXml += '<a:endParaRPr lang="' + (slideObj.options.lang || 'en-US') + '" dirty="0"/>' // NOTE: Added 20180101 to address PPT-2007 issues
	}
	strSlideXml += '</a:p>'

	// STEP 6: Close the textBody
	strSlideXml += tagClose

	// LAST: Return XML
	return strSlideXml
}

function genXmlParagraphProperties(textObj, isDefault) {
	var strXmlBullet = '',
		strXmlLnSpc = '',
		strXmlParaSpc = '',
		paraPropXmlCore = ''
	var bulletLvl0Margin = 342900
	var tag = isDefault ? 'a:lvl1pPr' : 'a:pPr'

	var paragraphPropXml = '<' + tag + (textObj.options.rtlMode ? ' rtl="1" ' : '')

	// A: Build paragraphProperties
	{
		// OPTION: align
		if (textObj.options.align) {
			switch (textObj.options.align) {
				case 'l':
				case 'left':
					paragraphPropXml += ' algn="l"'
					break
				case 'r':
				case 'right':
					paragraphPropXml += ' algn="r"'
					break
				case 'c':
				case 'ctr':
				case 'center':
					paragraphPropXml += ' algn="ctr"'
					break
				case 'justify':
					paragraphPropXml += ' algn="just"'
					break
			}
		}

		if (textObj.options.lineSpacing) {
			strXmlLnSpc = '<a:lnSpc><a:spcPts val="' + textObj.options.lineSpacing + '00"/></a:lnSpc>'
		}

		// OPTION: indent
		if (textObj.options.indentLevel && !isNaN(Number(textObj.options.indentLevel)) && textObj.options.indentLevel > 0) {
			paragraphPropXml += ' lvl="' + textObj.options.indentLevel + '"'
		}

		// OPTION: Paragraph Spacing: Before/After
		if (textObj.options.paraSpaceBefore && !isNaN(Number(textObj.options.paraSpaceBefore)) && textObj.options.paraSpaceBefore > 0) {
			strXmlParaSpc += '<a:spcBef><a:spcPts val="' + textObj.options.paraSpaceBefore * 100 + '"/></a:spcBef>'
		}
		if (textObj.options.paraSpaceAfter && !isNaN(Number(textObj.options.paraSpaceAfter)) && textObj.options.paraSpaceAfter > 0) {
			strXmlParaSpc += '<a:spcAft><a:spcPts val="' + textObj.options.paraSpaceAfter * 100 + '"/></a:spcAft>'
		}

		// Set core XML for use below
		paraPropXmlCore = paragraphPropXml

		// OPTION: bullet
		// NOTE: OOXML uses the unicode character set for Bullets
		// EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
		if (typeof textObj.options.bullet === 'object') {
			if (textObj.options.bullet.type) {
				if (textObj.options.bullet.type.toString().toLowerCase() == 'number') {
					paragraphPropXml +=
						' marL="' +
						(textObj.options.indentLevel && textObj.options.indentLevel > 0
							? bulletLvl0Margin + bulletLvl0Margin * textObj.options.indentLevel
							: bulletLvl0Margin) +
						'" indent="-' +
						bulletLvl0Margin +
						'"'
					strXmlBullet = '<a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>'
				}
			} else if (textObj.options.bullet.code) {
				var bulletCode = '&#x' + textObj.options.bullet.code + ';'

				// Check value for hex-ness (s/b 4 char hex)
				if (/^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.code) == false) {
					console.warn('Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!')
					bulletCode = BULLET_TYPES['DEFAULT']
				}

				paragraphPropXml +=
					' marL="' +
					(textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin + bulletLvl0Margin * textObj.options.indentLevel : bulletLvl0Margin) +
					'" indent="-' +
					bulletLvl0Margin +
					'"'
				strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>'
			}
		} else if (textObj.options.bullet == true) {
			paragraphPropXml +=
				' marL="' +
				(textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin + bulletLvl0Margin * textObj.options.indentLevel : bulletLvl0Margin) +
				'" indent="-' +
				bulletLvl0Margin +
				'"'
			strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + BULLET_TYPES['DEFAULT'] + '"/>'
		} else {
			strXmlBullet = '<a:buNone/>'
		}

		// Close Paragraph-Properties --------------------
		// IMPORTANT: strXmlLnSpc, strXmlParaSpc, and strXmlBullet require strict ordering.
		//            anything out of order is ignored. (PPT-Online, PPT for Mac)
		paragraphPropXml += '>' + strXmlLnSpc + strXmlParaSpc + strXmlBullet
		if (isDefault) {
			paragraphPropXml += genXmlTextRunProperties(textObj.options, true)
		}
		paragraphPropXml += '</' + tag + '>'
	}

	return paragraphPropXml
}

function genXmlTextRunProperties(opts, isDefault) {
	var runProps = ''
	var runPropsTag = isDefault ? 'a:defRPr' : 'a:rPr'

	// BEGIN runProperties
	runProps += '<' + runPropsTag + ' lang="' + (opts.lang ? opts.lang : 'en-US') + '" ' + (opts.lang ? ' altLang="en-US"' : '')
	runProps += opts.bold ? ' b="1"' : ''
	runProps += opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : '' // NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
	runProps += opts.italic ? ' i="1"' : ''
	runProps += opts.strike ? ' strike="sngStrike"' : ''
	runProps += opts.underline || opts.hyperlink ? ' u="sng"' : ''
	runProps += opts.subscript ? ' baseline="-40000"' : opts.superscript ? ' baseline="30000"' : ''
	runProps += opts.charSpacing ? ' spc="' + opts.charSpacing * 100 + '" kern="0"' : '' // IMPORTANT: Also disable kerning; otherwise text won't actually expand
	runProps += ' dirty="0" smtClean="0">'
	// Color / Font / Outline are children of <a:rPr>, so add them now before closing the runProperties tag
	if (opts.color || opts.fontFace || opts.outline) {
		if (opts.outline && typeof opts.outline === 'object') {
			runProps += '<a:ln w="' + Math.round((opts.outline.size || 0.75) * ONEPT) + '">' + genXmlColorSelection(opts.outline.color || 'FFFFFF') + '</a:ln>'
		}
		if (opts.color) runProps += genXmlColorSelection(opts.color)
		if (opts.fontFace) {
			// NOTE: 'cs' = Complex Script, 'ea' = East Asian (use -120 instead of 0 - see Issue #174); ea must come first (see Issue #174)
			runProps +=
				'<a:latin typeface="' +
				opts.fontFace +
				'" pitchFamily="34" charset="0" />' +
				'<a:ea typeface="' +
				opts.fontFace +
				'" pitchFamily="34" charset="-122" />' +
				'<a:cs typeface="' +
				opts.fontFace +
				'" pitchFamily="34" charset="-120" />'
		}
	}

	// Hyperlink support
	if (opts.hyperlink) {
		if (typeof opts.hyperlink !== 'object') console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` ")
		else if (!opts.hyperlink.url && !opts.hyperlink.slide) console.log("ERROR: 'hyperlink requires either `url` or `slide`'")
		else if (opts.hyperlink.url) {
			// FIXME-20170410: FUTURE-FEATURE: color (link is always blue in Keynote and PPT online, so usual text run above isnt honored for links..?)
			//runProps += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010! (Issue#74)
			runProps +=
				'<a:hlinkClick r:id="rId' +
				opts.hyperlink.rId +
				'" invalidUrl="" action="" tgtFrame="" tooltip="' +
				(opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') +
				'" history="1" highlightClick="0" endSnd="0" />'
		} else if (opts.hyperlink.slide) {
			runProps +=
				'<a:hlinkClick r:id="rId' +
				opts.hyperlink.rId +
				'" action="ppaction://hlinksldjump" tooltip="' +
				(opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') +
				'" />'
		}
	}

	// END runProperties
	runProps += '</' + runPropsTag + '>'

	return runProps
}

/**
* DESC: Builds <a:r></a:r> text runs for <a:p> paragraphs in textBody
* EX:
<a:r>
  <a:rPr lang="en-US" sz="2800" dirty="0" smtClean="0">
	<a:solidFill>
	  <a:srgbClr val="00FF00">
	  </a:srgbClr>
	</a:solidFill>
	<a:latin typeface="Courier New" pitchFamily="34" charset="0"/>
  </a:rPr>
  <a:t>Misc font/color, size = 28</a:t>
</a:r>
*/
function genXmlTextRun(opts, inStrText) {
	var xmlTextRun = ''
	var paraProp = ''
	var parsedText

	// ADD runProperties
	var startInfo = genXmlTextRunProperties(opts, false)

	// LINE-BREAKS/MULTI-LINE: Split text into multi-p:
	parsedText = inStrText.split(CRLF)
	if (parsedText.length > 1) {
		var outTextData = ''
		for (var i = 0, total_size_i = parsedText.length; i < total_size_i; i++) {
			outTextData += '<a:r>' + startInfo + '<a:t>' + encodeXmlEntities(parsedText[i])
			// Stop/Start <p>aragraph as long as there is more lines ahead (otherwise its closed at the end of this function)
			if (i + 1 < total_size_i) outTextData += (opts.breakLine ? CRLF : '') + '</a:t></a:r>'
		}
		xmlTextRun = outTextData
	} else {
		// Handle cases where addText `text` was an array of objects - if a text object doesnt contain a '\n' it still need alignment!
		// The first pPr-align is done in makeXml - use line countr to ensure we only add subsequently as needed
		xmlTextRun = (opts.align && opts.lineIdx > 0 ? paraProp : '') + '<a:r>' + startInfo + '<a:t>' + encodeXmlEntities(inStrText)
	}

	// Return paragraph with text run
	return xmlTextRun + '</a:t></a:r>'
}

/**
 * DESC: Builds <a:bodyPr></a:bodyPr> tag
 */
function genXmlBodyProperties(objOptions) {
	var bodyProperties = '<a:bodyPr'

	if (objOptions && objOptions.bodyProp) {
		// A: Enable or disable textwrapping none or square:
		objOptions.bodyProp.wrap ? (bodyProperties += ' wrap="' + objOptions.bodyProp.wrap + '" rtlCol="0"') : (bodyProperties += ' wrap="square" rtlCol="0"')

		// B: Set anchorPoints:
		if (objOptions.bodyProp.anchor) bodyProperties += ' anchor="' + objOptions.bodyProp.anchor + '"' // VALS: [t,ctr,b]
		if (objOptions.bodyProp.vert) bodyProperties += ' vert="' + objOptions.bodyProp.vert + '"' // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]

		// C: Textbox margins [padding]:
		if (objOptions.bodyProp.bIns || objOptions.bodyProp.bIns == 0) bodyProperties += ' bIns="' + objOptions.bodyProp.bIns + '"'
		if (objOptions.bodyProp.lIns || objOptions.bodyProp.lIns == 0) bodyProperties += ' lIns="' + objOptions.bodyProp.lIns + '"'
		if (objOptions.bodyProp.rIns || objOptions.bodyProp.rIns == 0) bodyProperties += ' rIns="' + objOptions.bodyProp.rIns + '"'
		if (objOptions.bodyProp.tIns || objOptions.bodyProp.tIns == 0) bodyProperties += ' tIns="' + objOptions.bodyProp.tIns + '"'

		// D: Close <a:bodyPr element
		bodyProperties += '>'

		// E: NEW: Add autofit type tags
		if (objOptions.shrinkText) bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000" />' // MS-PPT > Format Shape > Text Options: "Shrink text on overflow"
		// MS-PPT > Format Shape > Text Options: "Resize shape to fit text" [spAutoFit]
		// NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
		bodyProperties += objOptions.bodyProp.autoFit !== false ? '<a:spAutoFit/>' : ''

		// LAST: Close bodyProp
		bodyProperties += '</a:bodyPr>'
	} else {
		// DEFAULT:
		bodyProperties += ' wrap="square" rtlCol="0">'
		bodyProperties += '</a:bodyPr>'
	}

	// LAST: Return Close bodyProp
	return objOptions.isTableCell ? '<a:bodyPr/>' : bodyProperties
}

function genXmlColorSelection(color_info, back_info?: string) {
	var colorVal
	var fillType = 'solid'
	var internalElements = ''
	var outText = ''

	if (back_info && typeof back_info === 'string') {
		outText += '<p:bg><p:bgPr>'
		outText += genXmlColorSelection(back_info.replace('#', ''))
		outText += '<a:effectLst/>'
		outText += '</p:bgPr></p:bg>'
	}

	if (color_info) {
		if (typeof color_info == 'string') colorVal = color_info
		else {
			if (color_info.type) fillType = color_info.type
			if (color_info.color) colorVal = color_info.color
			if (color_info.alpha) internalElements += '<a:alpha val="' + (100 - color_info.alpha) + '000"/>'
		}

		switch (fillType) {
			case 'solid':
				outText += '<a:solidFill>' + createColorElement(colorVal, internalElements) + '</a:solidFill>'
				break
		}
	}

	return outText
}

function genXmlPlaceholder(placeholderObj) {
	var strXml = ''

	if (placeholderObj) {
		var placeholderIdx = placeholderObj.options && placeholderObj.options.placeholderIdx ? placeholderObj.options.placeholderIdx : ''
		var placeholderType = placeholderObj.options && placeholderObj.options.placeholderType ? placeholderObj.options.placeholderType : ''

		strXml +=
			'<p:ph' +
			(placeholderIdx ? ' idx="' + placeholderIdx + '"' : '') +
			(placeholderType && PLACEHOLDER_TYPES[placeholderType] ? ' type="' + PLACEHOLDER_TYPES[placeholderType] + '"' : '') +
			(placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : '') +
			'/>'
	}
	return strXml
}

// XML-GEN: First 6 functions create the base /ppt files

export function makeXmlContTypes(slides: Array<ISlide>, slideLayouts, masterSlide?): string {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
	strXml += ' <Default Extension="xml" ContentType="application/xml"/>'
	strXml += ' <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
	strXml += ' <Default Extension="jpeg" ContentType="image/jpeg"/>'
	strXml += ' <Default Extension="jpg" ContentType="image/jpg"/>'

	// STEP 1: Add standard/any media types used in Presenation
	strXml += ' <Default Extension="png" ContentType="image/png"/>'
	strXml += ' <Default Extension="gif" ContentType="image/gif"/>'
	strXml += ' <Default Extension="m4v" ContentType="video/mp4"/>' // NOTE: Hard-Code this extension as it wont be created in loop below (as extn != type)
	strXml += ' <Default Extension="mp4" ContentType="video/mp4"/>' // NOTE: Hard-Code this extension as it wont be created in loop below (as extn != type)
	slides.forEach(slide => {
		;(slide.relsMedia || []).forEach(rel => {
			if (rel.type != 'image' && rel.type != 'online' && rel.type != 'chart' && rel.extn != 'm4v' && strXml.indexOf(rel.type) == -1) {
				strXml += ' <Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>'
			}
		})
	})
	strXml += ' <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
	strXml += ' <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'

	// STEP 2: Add presentation and slide master(s)/slide(s)
	strXml += ' <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
	strXml += ' <Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>'
	slides.forEach((slide, idx) => {
		strXml +=
			'<Override PartName="/ppt/slideMasters/slideMaster' +
			(idx + 1) +
			'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
		strXml += '<Override PartName="/ppt/slides/slide' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
		// add charts if any
		slide.rels.forEach(rel => {
			if (rel.type == 'chart') {
				strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
			}
		})
	})

	// STEP 3: Core PPT
	strXml += ' <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
	strXml += ' <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>'
	strXml += ' <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
	strXml += ' <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'

	// STEP 4: Add Slide Layouts
	slideLayouts.forEach((layout, idx) => {
		strXml +=
			'<Override PartName="/ppt/slideLayouts/slideLayout' +
			(idx + 1) +
			'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
		layout.rels.forEach(rel => {
			if (rel.type == 'chart') {
				strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
			}
		})
	})

	// STEP 5: Add notes slide(s)
	slides.forEach((_slide, idx) => {
		strXml +=
			' <Override PartName="/ppt/notesSlides/notesSlide' +
			(idx + 1) +
			'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
	})

	masterSlide.rels.forEach(rel => {
		if (rel.type == 'chart') {
			strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
		}
		if (rel.type != 'image' && rel.type != 'online' && rel.type != 'chart' && rel.extn != 'm4v' && strXml.indexOf(rel.type) == -1)
			strXml += ' <Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>'
	})

	// STEP 5: Finish XML (Resume core)
	strXml += ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
	strXml += ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
	strXml += '</Types>'

	return strXml
}

export function makeXmlRootRels() {
	var strXml =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		CRLF +
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
		'  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
		'  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
		'  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
		'</Relationships>'
	return strXml
}

export function makeXmlApp(slides: Array<ISlide>, company: string): string {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml +=
		'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
	strXml += '<TotalTime>0</TotalTime>'
	strXml += '<Words>0</Words>'
	strXml += '<Application>Microsoft Office PowerPoint</Application>'
	strXml += '<PresentationFormat>On-screen Show</PresentationFormat>'
	strXml += '<Paragraphs>0</Paragraphs>'
	strXml += '<Slides>' + slides.length + '</Slides>'
	strXml += '<Notes>' + slides.length + '</Notes>'
	strXml += '<HiddenSlides>0</HiddenSlides>'
	strXml += '<MMClips>0</MMClips>'
	strXml += '<ScaleCrop>false</ScaleCrop>'
	strXml += '<HeadingPairs>'
	strXml += '  <vt:vector size="4" baseType="variant">'
	strXml += '    <vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>'
	strXml += '    <vt:variant><vt:i4>1</vt:i4></vt:variant>'
	strXml += '    <vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>'
	strXml += '    <vt:variant><vt:i4>' + slides.length + '</vt:i4></vt:variant>'
	strXml += '  </vt:vector>'
	strXml += '</HeadingPairs>'
	strXml += '<TitlesOfParts>'
	strXml += '<vt:vector size="' + (slides.length + 1) + '" baseType="lpstr">'
	strXml += '<vt:lpstr>Office Theme</vt:lpstr>'
	slides.forEach((_slideObj, idx) => {
		strXml += '<vt:lpstr>Slide ' + (idx + 1) + '</vt:lpstr>'
	})
	strXml += '</vt:vector>'
	strXml += '</TitlesOfParts>'
	strXml += '<Company>' + company + '</Company>'
	strXml += '<LinksUpToDate>false</LinksUpToDate>'
	strXml += '<SharedDoc>false</SharedDoc>'
	strXml += '<HyperlinksChanged>false</HyperlinksChanged>'
	strXml += '<AppVersion>15.0000</AppVersion>'
	strXml += '</Properties>'

	return strXml
}

export function makeXmlCore(title: string, subject: string, author: string, revision: string): string {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml +=
		'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
	strXml += '<dc:title>' + encodeXmlEntities(title) + '</dc:title>'
	strXml += '<dc:subject>' + encodeXmlEntities(subject) + '</dc:subject>'
	strXml += '<dc:creator>' + encodeXmlEntities(author) + '</dc:creator>'
	strXml += '<cp:lastModifiedBy>' + encodeXmlEntities(author) + '</cp:lastModifiedBy>'
	strXml += '<cp:revision>' + revision + '</cp:revision>'
	strXml += '<dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:created>'
	strXml += '<dcterms:modified xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:modified>'
	strXml += '</cp:coreProperties>'
	return strXml
}

export function makeXmlPresentationRels(slides: Array<ISlide>): string {
	var intRelNum = 0
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
	strXml += '  <Relationship Id="rId1" Target="slideMasters/slideMaster1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>'
	intRelNum++
	for (var idx = 1; idx <= slides.length; idx++) {
		intRelNum++
		strXml +=
			'  <Relationship Id="rId' + intRelNum + '" Target="slides/slide' + idx + '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>'
	}
	intRelNum++
	strXml +=
		'  <Relationship Id="rId' +
		intRelNum +
		'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>' +
		'  <Relationship Id="rId' +
		(intRelNum + 1) +
		'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>' +
		'  <Relationship Id="rId' +
		(intRelNum + 2) +
		'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
		'  <Relationship Id="rId' +
		(intRelNum + 3) +
		'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>' +
		'  <Relationship Id="rId' +
		(intRelNum + 4) +
		'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>' +
		'</Relationships>'

	return strXml
}

// XML-GEN: Next 5 functions run 1-N times (once for each Slide)

/**
 * Generates XML for the slide file
 * @param {Object} objSlide - the slide object to transform into XML
 * @return {string} strXml - slide OOXML
 */
export function makeXmlSlide(objSlide: ISlide): string {
	// STEP 1: Generate slide XML - wrap generated text in full XML envelope
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml +=
		'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' +
		(objSlide && objSlide.hidden ? ' show="0"' : '') +
		'>'
	strXml += gObjPptxGenerators.slideObjectToXml(objSlide)
	strXml += '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
	strXml += '</p:sld>'

	// LAST: Return
	return strXml
}

export function getNotesFromSlide(objSlide: ISlide): string {
	var notesStr = ''
	objSlide.data.forEach(data => {
		if (data.type === 'notes') {
			notesStr += data.text
		}
	})
	return notesStr.replace(/\r*\n/g, CRLF)
}

export function makeXmlNotesSlide(objSlide: ISlide): string {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml +=
		'<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
	strXml +=
		'<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="" /><p:cNvGrpSpPr />' +
		'<p:nvPr /></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0" />' +
		'<a:ext cx="0" cy="0" /><a:chOff x="0" y="0" /><a:chExt cx="0" cy="0" />' +
		'</a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1" />' +
		'<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1" /></p:cNvSpPr>' +
		'<p:nvPr><p:ph type="sldImg" /></p:nvPr></p:nvSpPr><p:spPr />' +
		'</p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2" />' +
		'<p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr>' +
		'<p:ph type="body" idx="1" /></p:nvPr></p:nvSpPr><p:spPr />' +
		'<p:txBody><a:bodyPr /><a:lstStyle /><a:p><a:r>' +
		'<a:rPr lang="en-US" dirty="0" smtClean="0" /><a:t>' +
		encodeXmlEntities(getNotesFromSlide(objSlide)) +
		'</a:t></a:r><a:endParaRPr lang="en-US" dirty="0" /></a:p></p:txBody>' +
		'</p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3" />' +
		'<p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr>' +
		'<p:ph type="sldNum" sz="quarter" idx="10" /></p:nvPr></p:nvSpPr>' +
		'<p:spPr /><p:txBody><a:bodyPr /><a:lstStyle /><a:p>' +
		'<a:fld id="' +
		SLDNUMFLDID +
		'" type="slidenum">' +
		'<a:rPr lang="en-US" smtClean="0" /><a:t>' +
		objSlide.numb +
		'</a:t></a:fld><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp>' +
		'</p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">' +
		'<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991" />' +
		'</p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping /></p:clrMapOvr></p:notes>'
	return strXml
}

/**
 * Generates the XML layout resource from a layout object
 *
 * @param {ISlide} objSlideLayout - slide object that represents layout
 * @return {string} strXml - slide OOXML
 */
export function makeXmlLayout(objSlideLayout: ISlideLayout): string {
	// STEP 1: Generate slide XML - wrap generated text in full XML envelope
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml +=
		'<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1">'
	strXml += gObjPptxGenerators.slideObjectToXml(objSlideLayout as ISlide)
	strXml += '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
	strXml += '</p:sldLayout>'

	// LAST: Return
	return strXml
}

/**
 * Generates XML for the master file
 * @param {ISlide} objSlide - slide object that represents master slide layout
 * @param {ISlideLayout[]} slideLayouts - slide layouts
 * @return {string} strXml - slide OOXML
 */
export function makeXmlMaster(objSlide: ISlide, slideLayouts: Array<ISlideLayout>): string {
	// NOTE: Pass layouts as static rels because they are not referenced any time
	var layoutDefs = slideLayouts.map((_layoutDef, idx) => {
		return '<p:sldLayoutId id="' + (LAYOUT_IDX_SERIES_BASE + idx) + '" r:id="rId' + (objSlide.rels.length + idx + 1) + '"/>'
	})

	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml +=
		'<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
	strXml += gObjPptxGenerators.slideObjectToXml(objSlide)
	strXml +=
		'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
	strXml += '<p:sldLayoutIdLst>' + layoutDefs.join('') + '</p:sldLayoutIdLst>'
	strXml += '<p:hf sldNum="0" hdr="0" ftr="0" dt="0"/>'
	strXml +=
		'<p:txStyles>' +
		' <p:titleStyle>' +
		'  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>' +
		' </p:titleStyle>' +
		' <p:bodyStyle>' +
		'  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
		'  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
		'  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
		'  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
		'  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
		'  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
		'  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
		'  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
		'  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
		' </p:bodyStyle>' +
		' <p:otherStyle>' +
		'  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
		'  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
		'  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
		'  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
		'  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
		'  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
		'  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
		'  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
		'  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
		'  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
		' </p:otherStyle>' +
		'</p:txStyles>'
	strXml += '</p:sldMaster>'

	// LAST: Return
	return strXml
}

/**
 * Generate XML for Notes Master
 *
 * @returns {string} XML
 */
export function makeXmlNotesMaster(): string {
	return (
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		CRLF +
		'<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1" /></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="" /><p:cNvGrpSpPr /><p:nvPr /></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="0" cy="0" /><a:chOff x="0" y="0" /><a:chExt cx="0" cy="0" /></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="2971800" cy="458788" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" /><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200" /></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="dt" idx="1" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="0" /><a:ext cx="2971800" cy="458788" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" /><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200" /></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{5282F153-3F37-0F45-9E97-73ACFA13230C}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0" /><a:t>6/20/18</a:t></a:fld><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Image Placeholder 3" /><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1" /></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000" /><a:ext cx="5486400" cy="3086100" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom><a:noFill /><a:ln w="12700"><a:solidFill><a:prstClr val="black" /></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr" /><a:lstStyle /><a:p><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Notes Placeholder 4" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550" /><a:ext cx="5486400" cy="3600450" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" /><a:lstStyle /><a:p><a:pPr lvl="0" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Footer Placeholder 5" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="4" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="8685213" /><a:ext cx="2971800" cy="458787" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b" /><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200" /></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="7" name="Slide Number Placeholder 6" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="5" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="8685213" /><a:ext cx="2971800" cy="458787" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b" /><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200" /></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{CE5E9CC1-C706-0F49-92D6-E571CC5EEA8F}" type="slidenum"><a:rPr lang="en-US" smtClean="0" /><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991" /></p:ext></p:extLst></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink" /><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl9pPr></p:notesStyle></p:notesMaster>'
	)
}

/**
 * Generates XML string for a slide layout relation file.
 * @param {Number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @return {String} complete XML string ready to be saved as a file
 */
export function makeXmlSlideLayoutRel(layoutNumber: number, slideLayouts: Array<ISlideLayout>): string {
	return gObjPptxGenerators.slideObjectRelationsToXml(slideLayouts[layoutNumber - 1] as ISlide, [
		{
			target: '../slideMasters/slideMaster1.xml',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
		},
	])
}

/**
 * Generates XML string for a slide relation file.
 * @param {Number} slideNumber 1-indexed number of a layout that relations are generated for
 * @return {string} complete XML string ready to be saved as a file
 */
export function makeXmlSlideRel(slides: Array<ISlide>, slideLayouts: Array<ISlideLayout>, slideNumber: number): string {
	return gObjPptxGenerators.slideObjectRelationsToXml(slides[slideNumber - 1], [
		{
			target: '../slideLayouts/slideLayout' + getLayoutIdxForSlide(slides, slideLayouts, slideNumber) + '.xml',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
		},
		{
			target: '../notesSlides/notesSlide' + slideNumber + '.xml',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
		},
	])
}

/**
 * Generates XML string for a slide relation file.
 * @param {Number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {String} complete XML string ready to be saved as a file
 */
export function makeXmlNotesSlideRel(slideNumber: number): string {
	return (
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		CRLF +
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>' +
		'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide' +
		slideNumber +
		'.xml"/>' +
		'</Relationships>'
	)
}

/**
 * Generates XML string for the master file.
 * @param {ISlide} `masterSlideObject` - slide object
 * @return {String} complete XML string ready to be saved as a file
 */
export function makeXmlMasterRel(masterSlideObject: ISlide, slideLayouts: Array<ISlideLayout>): string {
	var defaultRels = slideLayouts.map((_layoutDef, idx) => {
		return { target: '../slideLayouts/slideLayout' + (idx + 1) + '.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' }
	})
	defaultRels.push({ target: '../theme/theme1.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' })

	return gObjPptxGenerators.slideObjectRelationsToXml(masterSlideObject, defaultRels)
}

export function makeXmlNotesMasterRel(): string {
	return (
		'<?xml version="1.0" encoding="UTF-8"?>' +
		CRLF +
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
		'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>' +
		'</Relationships>'
	)
}

/**
 * For the passed slide number, resolves name of a layout that is used for.
 * @param {ISlide[]} `slides` - Array of slides
 * @param {Number} `slideLayouts`
 * @param {Number} slideNumber
 * @return {Number} slide number
 */
function getLayoutIdxForSlide(slides: Array<ISlide>, slideLayouts: Array<ISlideLayout>, slideNumber: number): number {
	var layoutName = slides[slideNumber - 1].layoutName

	for (var i = 0; i < slideLayouts.length; i++) {
		if (slideLayouts[i].name === layoutName) {
			return i + 1
		}
	}

	// IMPORTANT: Return 1 (for `slideLayout1.xml`) when no def is found
	// So all objects are in Layout1 and every slide that references it uses this layout.
	return 1
}

// XML-GEN: Last 5 functions create root /ppt files

export function makeXmlTheme() {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml +=
		'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">\
					<a:themeElements>\
					  <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>\
					  <a:dk2><a:srgbClr val="A7A7A7"/></a:dk2>\
					  <a:lt2><a:srgbClr val="535353"/></a:lt2>\
					  <a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>\
					  <a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>\
					  <a:fontScheme name="Office">\
					  <a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont>\
					  <a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/>\
					  </a:minorFont></a:fontScheme>\
					  <a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>\
					</a:theme>'
	return strXml
}

/**
* Create the `ppt/presentation.xml` file XML
* @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
* @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
* @param `slides` {Array<ISlide>} presentation slides
* @param `pptLayout` {ISlideLayout} presentation layout
*/
export function makeXmlPresentation(slides: Array<ISlide>, pptLayout: ILayout) {
	var strXml =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		CRLF +
		'<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
		'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
		'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
		(this._rtlMode ? 'rtl="1"' : '') +
		' saveSubsetFonts="1" autoCompressPictures="0">'

	// STEP 1: Build SLIDE master list
	strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
	strXml += '<p:sldIdLst>'
	for (var idx = 0; idx < slides.length; idx++) {
		strXml += '<p:sldId id="' + (idx + 256) + '" r:id="rId' + (idx + 2) + '"/>'
	}
	strXml += '</p:sldIdLst>'

	// Step 2: Add NOTES master list
	strXml += '<p:notesMasterIdLst><p:notesMasterId r:id="rId' + (slides.length + 2 + 4) + '"/></p:notesMasterIdLst>' // length+2+4 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic)

	// STEP 3: Build SLIDE text styles
	strXml +=
		'<p:sldSz cx="' +
		pptLayout.width +
		'" cy="' +
		pptLayout.height +
		'" type="' +
		pptLayout.name +
		'"/>' +
		'<p:notesSz cx="' +
		pptLayout.height +
		'" cy="' +
		pptLayout.width +
		'"/>' +
		'<p:defaultTextStyle>'
	;+'  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>'
	for (let idx = 1; idx < 10; idx++) {
		let intCurPos = 0
		strXml +=
			'  <a:lvl' +
			idx +
			'pPr marL="' +
			intCurPos +
			'" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">' +
			'    <a:defRPr sz="1800" kern="1200">' +
			'      <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
			'      <a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>' +
			'    </a:defRPr>' +
			'  </a:lvl' +
			idx +
			'pPr>'
		intCurPos += 457200
	}
	strXml += '</p:defaultTextStyle>'
	strXml += '</p:presentation>'

	return strXml
}

export function makeXmlPresProps() {
	var strXml =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		CRLF +
		'<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'

	return strXml
}

export function makeXmlTableStyles() {
	// SEE: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
	var strXml =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		CRLF +
		'<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'
	return strXml
}

export function makeXmlViewProps() {
	var strXml =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		CRLF +
		'<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
		'<p:normalViewPr><p:restoredLeft sz="15610" /><p:restoredTop sz="94613" /></p:normalViewPr>' +
		'<p:slideViewPr>' +
		'  <p:cSldViewPr snapToGrid="0" snapToObjects="1">' +
		'    <p:cViewPr varScale="1"><p:scale><a:sx n="119" d="100" /><a:sy n="119" d="100" /></p:scale><p:origin x="312" y="184" /></p:cViewPr>' +
		'    <p:guideLst />' +
		'  </p:cSldViewPr>' +
		'</p:slideViewPr>' +
		'<p:notesTextViewPr>' +
		'  <p:cViewPr><p:scale><a:sx n="1" d="1" /><a:sy n="1" d="1" /></p:scale><p:origin x="0" y="0" /></p:cViewPr>' +
		'</p:notesTextViewPr>' +
		'<p:gridSpacing cx="76200" cy="76200" />' +
		'</p:viewPr>'
	return strXml
}

/**
 * Create Grid Line Element
 * @param {Object} glOpts {size, color, style}
 * @param {Object} defaults {size, color, style}
 * @param {String} type "major"(default) | "minor"
 */
function createGridLineElement(glOpts: IChartOpts['catGridLine']) {
	let strXml = '<MajorGridlines>'
	strXml += ' <c:spPr>'
	strXml += '  <a:ln w="' + Math.round((glOpts.size || DEF_CHART_GRIDLINE.size) * ONEPT) + '" cap="flat">'
	strXml += '  <a:solidFill><a:srgbClr val="' + (glOpts.color || DEF_CHART_GRIDLINE.color) + '"/></a:solidFill>' // should accept scheme colors as implemented in PR 135
	strXml += '   <a:prstDash val="' + (glOpts.style || DEF_CHART_GRIDLINE.style) + '"/><a:round/>'
	strXml += '  </a:ln>'
	strXml += ' </c:spPr>'
	strXml += '</c:MajorGridlines>'

	return strXml
}

/**
 * DESC: Calc and return excel column name (eg: 'A2')
 */
function getExcelColName(length) {
	var strName = ''

	if (length <= 26) {
		strName = LETTERS[length]
	} else {
		strName += LETTERS[Math.floor(length / LETTERS.length) - 1]
		strName += LETTERS[length % LETTERS.length]
	}

	return strName
}

/**
 * DESC: Depending on the passed color string, creates either `a:schemeClr` (when scheme color) or `a:srgbClr` (when hexa representation).
 * color (string): hexa representation (eg. "FFFF00") or a scheme color constant (eg. colors.ACCENT1)
 * innerElements (optional string): Additional elements that adjust the color and are enclosed by the color element.
 */
function createColorElement(colorStr: string, innerElements?: string) {
	var isHexaRgb = REGEX_HEX_COLOR.test(colorStr)
	if (!isHexaRgb && !SCHEME_COLOR_NAMES[colorStr]) {
		console.warn('"' + colorStr + '" is not a valid scheme color or hexa RGB! "' + DEF_FONT_COLOR + '" is used as a fallback. Pass 6-digit RGB or `pptx.colors` values')
		colorStr = DEF_FONT_COLOR
	}
	var tagName = isHexaRgb ? 'srgbClr' : 'schemeClr'
	var colorAttr = ' val="' + colorStr + '"'
	return innerElements ? '<a:' + tagName + colorAttr + '>' + innerElements + '</a:' + tagName + '>' : '<a:' + tagName + colorAttr + ' />'
}

/**
 * NOTE: Used by both: text and lineChart
 * Creates `a:innerShdw` or `a:outerShdw` depending on pass options `opts`.
 * @param {Object} opts optional shadow properties
 * @param {Object} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 *	{ type: 'outer', blur: 3, offset: (23000 / 12700), angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true };
 */
function createShadowElement(options: IShadowOpts, defaults: object) {
	///if (options === 'none') {
	if (options === null) {
		return '<a:effectLst/>'
	}
	var strXml = '<a:effectLst>',
		opts = getMix(defaults, options),
		type = opts['type'] || 'outer',
		blur = opts['blur'] * ONEPT,
		offset = opts['offset'] * ONEPT,
		angle = opts['angle'] * 60000,
		color = opts['color'],
		opacity = opts['opacity'] * 100000,
		rotateWithShape = opts['rotateWithShape'] ? 1 : 0

	strXml += '<a:' + type + 'Shdw sx="100000" sy="100000" kx="0" ky="0"  algn="bl" blurRad="' + blur + '" '
	strXml += 'rotWithShape="' + +rotateWithShape + '"'
	strXml += ' dist="' + offset + '" dir="' + angle + '">'
	strXml += '<a:srgbClr val="' + color + '">' // TODO: should accept scheme colors implemented in Issue #135
	strXml += '<a:alpha val="' + opacity + '"/></a:srgbClr>'
	strXml += '</a:' + type + 'Shdw>'
	strXml += '</a:effectLst>'

	return strXml
}

/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {IShadowOpts} shadowOpts
 */
function correctShadowOptions(shadowOpts: IShadowOpts) {
	if (!shadowOpts || shadowOpts === null) return

	// OPT: `type`
	if (shadowOpts.type != 'outer' && shadowOpts.type != 'inner') {
		console.warn('Warning: shadow.type options are `outer` or `inner`.')
		shadowOpts.type = 'outer'
	}

	// OPT: `angle`
	if (shadowOpts.angle) {
		// A: REALITY-CHECK
		if (isNaN(Number(shadowOpts.angle)) || shadowOpts.angle < 0 || shadowOpts.angle > 359) {
			console.warn('Warning: shadow.angle can only be 0-359')
			shadowOpts.angle = 270
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		shadowOpts.angle = Math.round(Number(shadowOpts.angle))
	}

	// OPT: `opacity`
	if (shadowOpts.opacity) {
		// A: REALITY-CHECK
		if (isNaN(Number(shadowOpts.opacity)) || shadowOpts.opacity < 0 || shadowOpts.opacity > 1) {
			console.warn('Warning: shadow.opacity can only be 0-1')
			shadowOpts.opacity = 0.75
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		shadowOpts.opacity = Number(shadowOpts.opacity)
	}
}

function correctGridLineOptions(glOpts) {
	if (!glOpts || glOpts === 'none') return
	if (glOpts.size !== undefined && (isNaN(Number(glOpts.size)) || glOpts.size <= 0)) {
		console.warn('Warning: chart.gridLine.size must be greater than 0.')
		delete glOpts.size // delete prop to used defaults
	}
	if (glOpts.style && ['solid', 'dash', 'dot'].indexOf(glOpts.style) < 0) {
		console.warn('Warning: chart.gridLine.style options: `solid`, `dash`, `dot`.')
		delete glOpts.style
	}
}

function getShapeInfo(shapeName) {
	if (!shapeName) return gObjPptxShapes.RECTANGLE

	if (typeof shapeName == 'object' && shapeName.name && shapeName.displayName && shapeName.avLst) return shapeName

	if (gObjPptxShapes[shapeName]) return gObjPptxShapes[shapeName]

	var objShape = Object.keys(gObjPptxShapes).filter((key: string) => {
		return gObjPptxShapes[key].name == shapeName || gObjPptxShapes[key].displayName
	})[0]
	if (typeof objShape !== 'undefined' && objShape != null) return objShape

	return gObjPptxShapes.RECTANGLE
}

export function createHyperlinkRels(slides: Array<ISlide>, inText, slideRels) {
	var arrTextObjects = []

	// Only text objects can have hyperlinks, so return if this is plain text/number
	if (typeof inText === 'string' || typeof inText === 'number') return
	// IMPORTANT: Check for isArray before typeof=object, or we'll exhaust recursion!
	else if (Array.isArray(inText)) arrTextObjects = inText
	else if (typeof inText === 'object') arrTextObjects = [inText]

	arrTextObjects.forEach(text => {
		// `text` can be an array of other `text` objects (table cell word-level formatting), so use recursion
		if (Array.isArray(text)) createHyperlinkRels(slides, text, slideRels)
		else if (text && typeof text === 'object' && text.options && text.options.hyperlink && !text.options.hyperlink.rId) {
			if (typeof text.options.hyperlink !== 'object') console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:'https://github.com'}` ")
			else if (!text.options.hyperlink.url && !text.options.hyperlink.slide) console.log("ERROR: 'hyperlink requires either: `url` or `slide`'")
			else {
				var intRels = 0
				slides.forEach((slide, idx) => {
					intRels += slide.rels.length
				})
				var intRelId = intRels + 1

				slideRels.push({
					type: 'hyperlink',
					data: text.options.hyperlink.slide ? 'slide' : 'dummy',
					rId: intRelId,
					Target: text.options.hyperlink.url || text.options.hyperlink.slide,
				})

				text.options.hyperlink.rId = intRelId
			}
		}
	})
}
