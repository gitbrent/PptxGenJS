/**
 * PptxGenJS: Table Generation
 */

import { CRLF, DEF_FONT_SIZE, DEF_SLIDE_MARGIN_IN, EMU, LINEH_MODIFIER, ONEPT, SLIDE_OBJECT_TYPES } from './core-enums'
import PptxGenJS from './pptxgen'
import { ILayout, ISlideLayout, TableCell, ITableToSlidesCell, ITableToSlidesOpts, ITableRow, TableRowSlide } from './core-interfaces'
import { inch2Emu, rgbToHex } from './gen-utils'

/**
 * Break text paragraphs into lines based upon table column width (e.g.: Magic Happens Here(tm))
 * @param {TableCell} cell - table cell
 * @param {number} colWidth - table column width
 * @return {string[]} XML
 */
function parseTextToLines(cell: TableCell, colWidth: number): string[] {
	let CHAR = 2.2 + (cell.options && cell.options.autoPageCharWeight ? cell.options.autoPageCharWeight : 0) // Character Constant (An approximation of the Golden Ratio)
	let CPL = (colWidth * EMU) / (((cell.options && cell.options.fontSize) || DEF_FONT_SIZE) / CHAR) // Chars-Per-Line
	let arrLines = []
	let strCurrLine = ''

	// A: Allow a single space/whitespace as cell text (user-requested feature)
	if (cell.text && cell.text.toString().trim() == '') return [' ']

	// B: Remove leading/trailing spaces
	let inStr = (cell.text || '').toString().trim()

	// C: Build line array
	// FIXME: FIXME-3: change to `forEach`
	jQuery.each(inStr.split('\n'), (_idx, line) => {
		jQuery.each(line.split(' '), (_idx, word) => {
			if (strCurrLine.length + word.length + 1 < CPL) {
				strCurrLine += word + ' '
			} else {
				if (strCurrLine) arrLines.push(strCurrLine)
				strCurrLine = word + ' '
			}
		})

		// All words for this line have been exhausted, flush buffer to new line, clear line var
		if (strCurrLine) arrLines.push(strCurrLine.trim() + CRLF)
		strCurrLine = ''
	})

	// D: Remove trailing linebreak
	arrLines[arrLines.length - 1] = arrLines[arrLines.length - 1].trim()

	return arrLines
}

export function getSlidesForTableRows(
	tableRows: [ITableToSlidesCell[]?] = [],
	tabOpts: ITableToSlidesOpts = {},
	presLayout: ILayout,
	masterSlide: ISlideLayout
): TableRowSlide[] {
	let arrInchMargins = DEF_SLIDE_MARGIN_IN,
		emuTabCurrH = 0,
		emuSlideTabW = EMU * 1,
		emuSlideTabH = EMU * 1,
		numCols = 0,
		tableRowSlides = [
			{
				rows: [] as ITableRow[],
			},
		]

	if (tabOpts.verbose) {
		console.log(`-- VERBOSE MODE ----------------------------------`)
		console.log(`.. (PARAMETERS)`)
		console.log(`presLayout.height ......... = ${presLayout.height / EMU}`)
		console.log(`tabOpts.h ................. = ${tabOpts.h}`)
		console.log(`tabOpts.w ................. = ${tabOpts.w}`)
		console.log(`tabOpts.colW .............. = ${tabOpts.colW}`)
		console.log(`tabOpts.slideMargin ....... = ${tabOpts.slideMargin || ''}`)
		console.log(`.. (/PARAMETERS)`)
	}

	// STEP 1: Calculate margins, overall usable slide space
	{
		// Important: Use default size as zero cell margin is causing our tables to be too large and touch bottom of slide!
		if (!tabOpts.slideMargin && tabOpts.slideMargin != 0) tabOpts.slideMargin = DEF_SLIDE_MARGIN_IN[0]

		if (tabOpts.slideMargin || tabOpts.slideMargin == 0) {
			if (Array.isArray(tabOpts.slideMargin)) arrInchMargins = tabOpts.slideMargin
			else if (!isNaN(tabOpts.slideMargin)) arrInchMargins = [tabOpts.slideMargin, tabOpts.slideMargin, tabOpts.slideMargin, tabOpts.slideMargin]
		} else if (masterSlide && masterSlide.margin) {
			if (Array.isArray(masterSlide.margin)) arrInchMargins = masterSlide.margin
			else if (!isNaN(masterSlide.margin)) arrInchMargins = [masterSlide.margin, masterSlide.margin, masterSlide.margin, masterSlide.margin]
		}

		if (tabOpts.verbose) console.log('arrInchMargins ......... = ' + arrInchMargins.toString())
	}

	// STEP 2: Calculate number of columns
	{
		// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
		// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
		tableRows[0].forEach(cell => {
			if (!cell) cell = { type: SLIDE_OBJECT_TYPES.tablecell }
			let cellOpts = cell.options || null
			numCols += Number(cellOpts && cellOpts.colspan ? cellOpts.colspan : 1)
		})

		if (tabOpts.verbose) console.log('numCols ................ = ' + numCols)
	}

	// STEP 3: Calculate tabOpts.w if tabOpts.colW was provided
	if (!tabOpts.w && tabOpts.colW) {
		if (Array.isArray(tabOpts.colW))
			tabOpts.colW.forEach(val => {
				typeof tabOpts.w !== 'number' ? (tabOpts.w = 0 + val) : (tabOpts.w += val)
			})
		else {
			tabOpts.w = tabOpts.colW * numCols
		}
	}

	// STEP 4: Calculate usable space/table size (now that total usable space is known)
	{
		emuSlideTabW =
			typeof tabOpts.w === 'number'
				? inch2Emu(tabOpts.w)
				: presLayout.width - inch2Emu((typeof tabOpts.x === 'number' ? tabOpts.x : arrInchMargins[1]) + arrInchMargins[3])

		if (tabOpts.verbose) console.log('emuSlideTabW (in) ...... = ' + (emuSlideTabW / EMU).toFixed(1))
	}

	// STEP 5: Calculate column widths if not provided (emuSlideTabW will be used below to determine lines-per-col)
	if (!tabOpts.colW || !Array.isArray(tabOpts.colW)) {
		if (tabOpts.colW && !isNaN(Number(tabOpts.colW))) {
			let arrColW = []
			tableRows[0].forEach(() => {
				arrColW.push(tabOpts.colW)
			})
			tabOpts.colW = []
			arrColW.forEach(val => {
				if (Array.isArray(tabOpts.colW)) tabOpts.colW.push(val)
			})
		}
		// No column widths provided? Then distribute cols.
		else {
			tabOpts.colW = []
			for (var iCol = 0; iCol < numCols; iCol++) {
				tabOpts.colW.push(emuSlideTabW / EMU / numCols)
			}
		}
	}

	// STEP 6: **MAIN** Iterate over rows, add table content, create new slides as rows overflow
	tableRows.forEach((row, iRow) => {
		// A: Row variables
		let maxLineHeight = 0
		let linesRow: TableCell[] = []

		// B: Create new row in data model
		let currSlide = tableRowSlides[tableRowSlides.length - 1]
		let newRowSlide = []
		row.forEach(cell => {
			newRowSlide.push({
				type: SLIDE_OBJECT_TYPES.tablecell,
				text: '',
				options: cell.options,
			})
		})

		// C: Calc usable vertical space/table height. Set default value first, adjust below when necessary.
		emuSlideTabH = tabOpts.h && typeof tabOpts.h === 'number' ? tabOpts.h : presLayout.height - inch2Emu(arrInchMargins[0] + arrInchMargins[2])

		// D: RULE: Use margins for starting point after the initial Slide, not `opt.y` (ISSUE#43, ISSUE#47, ISSUE#48)
		if (tableRowSlides.length > 1 && typeof tabOpts.newSlideStartY === 'number') {
			console.log(`tabOpts.newSlideStartY = ${tabOpts.newSlideStartY}`)
			emuSlideTabH = tabOpts.h && typeof tabOpts.h === 'number' ? tabOpts.h : presLayout.height - inch2Emu(tabOpts.newSlideStartY + arrInchMargins[2])
		} else if (tableRowSlides.length > 1 && typeof tabOpts.y === 'number') {
			emuSlideTabH = presLayout.height - inch2Emu((tabOpts.y / EMU < arrInchMargins[0] ? tabOpts.y / EMU : arrInchMargins[0]) + arrInchMargins[2])
			// Use whichever is greater: area between margins or the table H provided (dont shrink usable area - the whole point of over-riding X on paging is to *increarse* usable space)
			if (typeof tabOpts.h === 'number' && emuSlideTabH < tabOpts.h) emuSlideTabH = tabOpts.h
		} else if (typeof tabOpts.h === 'number' && typeof tabOpts.y === 'number')
			emuSlideTabH = tabOpts.h ? tabOpts.h : presLayout.height - inch2Emu((tabOpts.y / EMU || arrInchMargins[0]) + arrInchMargins[2])
		//if (tabOpts.verbose) console.log(`- SLIDE [${tableRowSlides.length}]: emuSlideTabH .. = ${(emuSlideTabH / EMU).toFixed(1)}`)

		// E: **BUILD DATA SET** | Iterate over cells: split text into lines[], set `lineHeight`
		row.forEach((cell, iCell) => {
			let newCell: TableCell = {
				type: SLIDE_OBJECT_TYPES.tablecell,
				text: '',
				options: cell.options,
				lines: [],
				lineHeight: inch2Emu(
					((cell.options && cell.options.fontSize ? cell.options.fontSize : tabOpts.fontSize ? tabOpts.fontSize : DEF_FONT_SIZE) *
						(LINEH_MODIFIER + (tabOpts.autoPageLineWeight ? tabOpts.autoPageLineWeight : 0))) /
						100
				),
			}
			//if (tabOpts.verbose) console.log(`- CELL [${iCell}]: newCell.lineHeight ..... = ${(newCell.lineHeight / EMU).toFixed(2)}`)

			// 1: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
			if (newCell.options.rowspan) newCell.lineHeight = 0

			// 2: The parseTextToLines method uses `autoPageCharWeight`, so inherit from table options
			newCell.options.autoPageCharWeight = tabOpts.autoPageCharWeight ? tabOpts.autoPageCharWeight : null

			// 3: **MAIN** Parse cell contents into lines based upon col width, font, etc
			newCell.lines = parseTextToLines(cell, tabOpts.colW[iCell] / ONEPT)

			// 4: Add to array
			linesRow.push(newCell)
		})

		/* F: **BUILD/PAGE DATA SET**
		 * Add text one-line-a-time to this row's cells until: lines are exhausted OR table height limit is hit
		 * Design: Building cells L-to-R/loop style wont work as one could be 100 lines and another 1 line.
		 * Therefore, build the whole row, 1-line-at-a-time, spanning all columns.
		 * That way, when the vertical size limit is hit, all lines pick up where they need to on the subsequent slide.
		 */
		let emuMaxCellMargin = 0
		// TODO: FIXME: find largest cell margin top and bottom, this plus emuTabCurrH is our size!
		// FIXME: im only using cell[0] for now as a DEV/DEBUG value
		emuMaxCellMargin = row[0].options.margin
			? row[0].options.margin[0] * ONEPT + row[0].options.margin[2] * ONEPT
			: tabOpts.margin
			? tabOpts.margin[0] * ONEPT + tabOpts.margin[2] * ONEPT
			: 0
		// Add T/B cell margins to calc actual row height
		emuTabCurrH += emuMaxCellMargin

		// G: Only create a new row if there is room, otherwise, it'll be an empty row as "A:" below will create a new Slide before loop can populate this row
		if (emuTabCurrH + maxLineHeight <= emuSlideTabH) currSlide.rows.push(newRowSlide)

		if (tabOpts.verbose) console.log(`- SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: START...`)
		while (
			linesRow.filter(cell => {
				return cell.lines.length > 0
			}).length > 0
		) {
			// A: Add new Slide if there is no more space to fix 1 new line
			if (emuTabCurrH + maxLineHeight > emuSlideTabH) {
				if (tabOpts.verbose)
					console.log(
						`** NEW SLIDE CREATED *****************************************` +
							` (why?): ${(emuTabCurrH / EMU).toFixed(2)}+${(maxLineHeight / EMU).toFixed(2)} > ${emuSlideTabH / EMU}`
					)

				// 1: Add a new slide
				tableRowSlides.push({
					rows: [] as ITableRow[],
				})

				// 2: Add new row to new slide
				let currSlide = tableRowSlides[tableRowSlides.length - 1]
				let newRowSlide = []
				row.forEach(cell => {
					newRowSlide.push({
						type: SLIDE_OBJECT_TYPES.tablecell,
						text: '',
						options: cell.options,
					})
				})
				currSlide.rows.push(newRowSlide)

				// TODO: FIXME: add support for "addHeaderToEach"
				// add here, increase `emuTabCurrH`

				// 3: Reset current table height for new Slide
				emuTabCurrH = 0 // This row's emuRowH w/b added below
			}

			// B: Add a line of text to 1-N cells that still have `lines`
			linesRow.forEach((cell, idx) => {
				if (cell.lines.length > 0) {
					// 1
					let currSlide = tableRowSlides[tableRowSlides.length - 1]
					currSlide.rows[currSlide.rows.length - 1][idx].text +=
						(currSlide.rows[currSlide.rows.length - 1][idx].text.length > 0 && !RegExp(/\n$/g).test(currSlide.rows[currSlide.rows.length - 1][idx].text)
							? CRLF
							: ''
						).replace(/[\r\n]+$/g, CRLF) + cell.lines.shift()

					// 2
					if (cell.lineHeight > maxLineHeight) maxLineHeight = cell.lineHeight
				}
			})

			// C: Increase table height by one line height as 1-N cells below are
			emuTabCurrH += maxLineHeight
			if (tabOpts.verbose) console.log(`- SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: one line added ... emuTabCurrH = ${(emuTabCurrH / EMU).toFixed(2)}`)
		}

		if (tabOpts.verbose)
			console.log(
				`- SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: ...COMPLETE ...... emuTabCurrH = ${(emuTabCurrH / EMU).toFixed(2)} ( emuSlideTabH = ${(
					emuSlideTabH / EMU
				).toFixed(2)} )`
			)
	})

	if (tabOpts.verbose) {
		console.log(`\n|================================================|\n| FINAL: tableRowSlides.length = ${tableRowSlides.length}`)
		console.log(tableRowSlides)
		//console.log(JSON.stringify(tableRowSlides,null,2))
		console.log(`|================================================|\n\n`)
	}

	return tableRowSlides
}

/**
 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
 * @param {string} `tabEleId` - HTMLElementID of the table
 * @param {ITableToSlidesOpts} `inOpts` - array of options (e.g.: tabsize)
 */
export function genTableToSlides(pptx: PptxGenJS, tabEleId: string, options: ITableToSlidesOpts = {}, masterSlide: ISlideLayout) {
	let opts = options || {}
	opts.slideMargin = opts.slideMargin || opts.slideMargin == 0 ? opts.slideMargin : 0.5
	let emuSlideTabW = opts.w || pptx.presLayout.width
	let arrObjTabHeadRows: [ITableToSlidesCell[]?] = []
	let arrObjTabBodyRows: [ITableToSlidesCell[]?] = []
	let arrObjTabFootRows: [ITableToSlidesCell[]?] = []
	let arrColW: number[] = []
	let arrTabColW: number[] = []
	let arrInchMargins: [number, number, number, number] = [0.5, 0.5, 0.5, 0.5] // TRBL-style
	let arrTableParts = ['thead', 'tbody', 'tfoot']
	let intTabW = 0

	// REALITY-CHECK:
	if (!document.getElementById(tabEleId)) throw 'tableToSlides: Table ID "' + tabEleId + '" does not exist!'

	// Set margins
	if (masterSlide && masterSlide.margin) {
		if (Array.isArray(masterSlide.margin)) arrInchMargins = masterSlide.margin
		else if (!isNaN(masterSlide.margin)) arrInchMargins = [masterSlide.margin, masterSlide.margin, masterSlide.margin, masterSlide.margin]
		opts.slideMargin = arrInchMargins
	} else if (opts && opts.slideMargin) {
		if (Array.isArray(opts.slideMargin)) arrInchMargins = opts.slideMargin
		else if (!isNaN(opts.slideMargin)) arrInchMargins = [opts.slideMargin, opts.slideMargin, opts.slideMargin, opts.slideMargin]
	}
	emuSlideTabW = (opts.w ? inch2Emu(opts.w) : pptx.presLayout.width) - inch2Emu(arrInchMargins[1] + arrInchMargins[3])

	if (opts.verbose) console.log('-- VERBOSE MODE ----------------------------------')
	if (opts.verbose) console.log(`opts.h ................. = ${opts.h}`)
	if (opts.verbose) console.log(`opts.w ................. = ${opts.w}`)
	if (opts.verbose) console.log(`pptx.presLayout.width .. = ${pptx.presLayout.width / EMU}`)
	if (opts.verbose) console.log(`emuSlideTabW (in)....... = ${emuSlideTabW / EMU}`)

	// STEP 1: Grab table col widths
	// ATTN: `arrTableParts.forEach((part, _idx) => {` --> NO! CAREFUL! We need to break out of loop using "return false" - forEach break col sizing badly
	jQuery.each(arrTableParts, (_idx, part) => {
		if (jQuery('#' + tabEleId + ' > ' + part + ' > tr').length > 0) {
			jQuery('#' + tabEleId + ' > ' + part + ' > tr:first-child')
				.find('> th, > td')
				.each((idx, cell) => {
					// FIXME: This is a hack - guessing at col widths when colspan
					if (jQuery(cell).attr('colspan')) {
						for (var idx = 0; idx < Number(jQuery(cell).attr('colspan')); idx++) {
							arrTabColW.push(Math.round(jQuery(cell).outerWidth() / Number(jQuery(cell).attr('colspan'))))
						}
					} else {
						arrTabColW.push(jQuery(cell).outerWidth())
					}
				})
			return false // break out of .each loop
		}
	})
	arrTabColW.forEach((colW, _idx) => {
		intTabW += colW
	})

	// STEP 2: Calc/Set column widths by using same column width percent from HTML table
	arrTabColW.forEach((colW, idx) => {
		let intCalcWidth = Number(((Number(emuSlideTabW) * ((colW / intTabW) * 100)) / 100 / EMU).toFixed(2))
		let intMinWidth = jQuery('#' + tabEleId + ' thead tr:first-child th:nth-child(' + (idx + 1) + ')').data('pptx-min-width')
		let intSetWidth = jQuery('#' + tabEleId + ' thead tr:first-child th:nth-child(' + (idx + 1) + ')').data('pptx-width')
		arrColW.push(intSetWidth ? intSetWidth : intMinWidth > intCalcWidth ? intMinWidth : intCalcWidth)
	})
	if (opts.verbose) console.log(`arrColW ................ = ${arrColW.toString()}`)

	// STEP 3: Iterate over each table element and create data arrays (text and opts)
	// NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
	arrTableParts.forEach((part, _idx) => {
		jQuery('#' + tabEleId + ' > ' + part + ' > tr').each((_idx, row) => {
			let arrObjTabCells = []
			jQuery(row)
				.find('> th, > td')
				.each((_idx, cell) => {
					// A: Get RGB text/bkgd colors
					let arrRGB1 = []
					let arrRGB2 = []
					arrRGB1 = jQuery(cell)
						.css('color')
						.replace(/\s+/gi, '')
						.replace('rgba(', '')
						.replace('rgb(', '')
						.replace(')', '')
						.split(',')
					arrRGB2 = jQuery(cell)
						.css('background-color')
						.replace(/\s+/gi, '')
						.replace('rgba(', '')
						.replace('rgb(', '')
						.replace(')', '')
						.split(',')
					// ISSUE#57: jQuery default is this rgba value of below giving unstyled tables a black bkgd, so use white instead
					// (FYI: if cell has `background:#000000` jQuery returns 'rgb(0, 0, 0)', so this soln is pretty solid)
					if (jQuery(cell).css('background-color') == 'rgba(0, 0, 0, 0)' || jQuery(cell).css('background-color') == 'transparent') arrRGB2 = [255, 255, 255]

					// B: Create option object
					let cellOpts = {
						fontSize: jQuery(cell)
							.css('font-size')
							.replace(/[a-z]/gi, ''),
						bold: jQuery(cell).css('font-weight') == 'bold' || Number(jQuery(cell).css('font-weight')) >= 500 ? true : false,
						color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
						fill: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2])),
						align: null,
						border: null,
						margin: null,
						colspan: null,
						rowspan: null,
						valign: null,
					}
					if (['left', 'center', 'right', 'start', 'end'].indexOf(jQuery(cell).css('text-align')) > -1)
						cellOpts.align = jQuery(cell)
							.css('text-align')
							.replace('start', 'left')
							.replace('end', 'right')
					if (['top', 'middle', 'bottom'].indexOf(jQuery(cell).css('vertical-align')) > -1) cellOpts.valign = jQuery(cell).css('vertical-align')

					// C: Add padding [margin] (if any)
					// NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
					if (jQuery(cell).css('padding-left')) {
						cellOpts.margin = []
						jQuery.each(['padding-top', 'padding-right', 'padding-bottom', 'padding-left'], (_idx, val) => {
							cellOpts.margin.push(
								Math.round(
									Number(
										jQuery(cell)
											.css(val)
											.replace(/\D/gi, '')
									)
								)
							)
						})
					}

					// D: Add colspan/rowspan (if any)
					if (jQuery(cell).attr('colspan')) cellOpts.colspan = jQuery(cell).attr('colspan')
					if (jQuery(cell).attr('rowspan')) cellOpts.rowspan = jQuery(cell).attr('rowspan')

					// E: Add border (if any)
					if (
						jQuery(cell).css('border-top-width') ||
						jQuery(cell).css('border-right-width') ||
						jQuery(cell).css('border-bottom-width') ||
						jQuery(cell).css('border-left-width')
					) {
						cellOpts.border = []
						jQuery.each(['top', 'right', 'bottom', 'left'], (_idx, val) => {
							var intBorderW = Math.round(
								Number(
									jQuery(cell)
										.css('border-' + val + '-width')
										.replace('px', '')
								)
							)
							var arrRGB = []
							arrRGB = jQuery(cell)
								.css('border-' + val + '-color')
								.replace(/\s+/gi, '')
								.replace('rgba(', '')
								.replace('rgb(', '')
								.replace(')', '')
								.split(',')
							var strBorderC = rgbToHex(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]))
							cellOpts.border.push({ pt: intBorderW, color: strBorderC })
						})
					}

					// F: Massage cell text so we honor linebreak tag as a line break during line parsing
					let $cell2 = jQuery(cell).clone()
					$cell2.html(
						jQuery(cell)
							.html()
							.replace(/<br[^>]*>/gi, '\n')
					)

					// LAST: Add cell
					arrObjTabCells.push({
						text: $cell2.text().trim(),
						options: cellOpts,
					})
				})
			switch (part) {
				case 'thead':
					arrObjTabHeadRows.push(arrObjTabCells)
					break
				case 'tbody':
					arrObjTabBodyRows.push(arrObjTabCells)
					break
				case 'tfoot':
					arrObjTabFootRows.push(arrObjTabCells)
					break
				default:
			}
		})
	})

	// STEP 5: Break table into Slides as needed
	// Pass head-rows as there is an option to add to each table and the parse func needs this data to fulfill that option
	opts._arrObjTabHeadRows = arrObjTabHeadRows || null
	opts.colW = arrColW

	getSlidesForTableRows(arrObjTabHeadRows.concat(arrObjTabBodyRows).concat(arrObjTabFootRows) as [ITableToSlidesCell[]], opts, pptx.presLayout, masterSlide).forEach(
		(slide, idx) => {
			// A: Create new Slide
			let newSlide = pptx.addSlide(opts.masterSlideName || null)

			// B: DESIGN: Reset `y` to `newSlideStartY` or margin after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
			if (idx == 0) opts.y = opts.y || arrInchMargins[0]
			if (idx > 0) opts.y = opts.newSlideStartY || arrInchMargins[0]
			if (opts.verbose) console.log('opts.newSlideStartY:' + opts.newSlideStartY + ' / arrInchMargins[0]:' + arrInchMargins[0] + ' => opts.y = ' + opts.y)

			// C: Add table to Slide
			newSlide.addTable(slide.rows, { x: opts.x || arrInchMargins[3], y: opts.y, w: Number(emuSlideTabW) / EMU, colW: arrColW, autoPage: false })

			// D: Add any additional objects
			if (opts.addImage) newSlide.addImage({ path: opts.addImage.url, x: opts.addImage.x, y: opts.addImage.y, w: opts.addImage.w, h: opts.addImage.h })
			if (opts.addShape) newSlide.addShape(opts.addShape.shape, opts.addShape.opts || {})
			if (opts.addTable) newSlide.addTable(opts.addTable.rows, opts.addTable.opts || {})
			if (opts.addText) newSlide.addText(opts.addText.text, opts.addText.opts || {})
		}
	)
}
