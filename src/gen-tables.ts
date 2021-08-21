/**
 * PptxGenJS: Table Generation
 */

import { CRLF, DEF_FONT_SIZE, DEF_SLIDE_MARGIN_IN, EMU, LINEH_MODIFIER, ONEPT, SLIDE_OBJECT_TYPES } from './core-enums'
import { PresLayout, SlideLayout, TableCell, TableToSlidesProps, TableRow, TableRowSlide, TableCellProps, ISlideObject } from './core-interfaces'
import { inch2Emu, rgbToHex, valToPts } from './gen-utils'
import PptxGenJS from './pptxgen'

/**
 * Break cell text into lines based upon table column width (e.g.: Magic Happens Here(tm))
 * @param {TableCell} cell - table cell
 * @param {number} colWidth - table column width
 * @return {TableRow[]} - cell's text objects grouped into lines
 */
function parseTextToLines(cell: TableCell, colWidth: number, verbose?: boolean): TableCell[][] {
	const CHAR = 2.2 + (cell.options && cell.options.autoPageCharWeight ? cell.options.autoPageCharWeight : 0) // Character Constant (An approximation of the Golden Ratio)
	const CPL = (colWidth * EMU) / (((cell.options && cell.options.fontSize) || DEF_FONT_SIZE) / CHAR) // Chars-Per-Line
	let parsedLines: TableCell[][] = []
	let inputCells: TableCell[] = []
	let inputLines1: TableCell[][] = []
	let inputLines2: TableCell[][] = []

	/**
	 * EX INPUTS: `cell.text`
	 * - string....: "Account Name Column"
	 * - object....: { text:"Account Name Column" }
	 * - object[]..: [{ text:"Account Name", options:{ bold:true } }, { text:" Column" }]
	 * - object[]..: [{ text:"Account Name", options:{ breakLine:true } }, { text:"Input" }]
	 */

	/**
	 * EX OUTPUTS:
	 * - string....: [{ text:"Account Name Column" }]
	 * - object....: [{ text:"Account Name Column" }]
	 * - object[]..: [{ text:"Account Name", options:{ breakLine:true } }, { text:"Input" }]
	 * - object[]..: [{ text:"Account Name", options:{ breakLine:true } }, { text:"Input" }]
	 */

	// STEP 1: Ensure inputCells is an array of TableCells
	if (cell.text && cell.text.toString().trim().length === 0) {
		// Allow a single space/whitespace as cell text (user-requested feature)
		inputCells.push({ _type: SLIDE_OBJECT_TYPES.tablecell, text: ' ' })
	} else if (typeof cell.text === 'number' || typeof cell.text === 'string') {
		inputCells.push({ _type: SLIDE_OBJECT_TYPES.tablecell, text: (cell.text || '').toString().trim() })
	} else if (Array.isArray(cell.text)) {
		inputCells = cell.text
	}
	if (verbose) {
		console.log('[1/4] inputCells')
		inputCells.forEach((cell, idx) => console.log(`[${idx + 1}] cell: ${JSON.stringify(cell)}`))
		console.log('...............................................\n\n')
	}

	// STEP 2: Group table cells into lines based on "\n" or `breakLine` prop
	/**
	 * - EX: `[{ text:"Input Output" }, { text:"Extra" }]`							== 1 line
	 * - EX: `[{ text:"Input" }, { text:"Output", options:{ breakLine:true } }]`	== 1 line
	 * - EX: `[{ text:"Input\nOutput" }]`											== 2 lines
	 * - EX: `[{ text:"Input", options:{ breakLine:true } }, { text:"Output" }]`	== 2 lines
	 */
	let newLine: TableCell[] = []
	inputCells.forEach(cell => {
		// (this is always true, we just constructed them above, but we need to tell typescript b/c type is still string||Cell[])
		if (typeof cell.text === 'string') {
			if (cell.text.split('\n').length > 1) {
				cell.text.split('\n').forEach(textLine => {
					newLine.push({
						_type: SLIDE_OBJECT_TYPES.tablecell,
						text: textLine,
						options: cell.options,
					})
				})
			} else {
				newLine.push({
					_type: SLIDE_OBJECT_TYPES.tablecell,
					text: cell.text.trim(),
					options: cell.options,
				})
			}

			if (cell.options && cell.options.breakLine) {
				if (verbose) console.log(`inputCells: new line > ${JSON.stringify(newLine)}`)
				inputLines1.push(newLine)
				newLine = []
			}
		}

		// Flush buffer
		if (newLine.length > 0) inputLines1.push(newLine)
	})
	if (verbose) {
		console.log(`[2/4] inputLines1 (${inputLines1.length})`)
		inputLines1.forEach((line, idx) => console.log(`[${idx + 1}] line: ${JSON.stringify(line)}`))
		console.log('...............................................\n\n')
	}

	// STEP 3: Tokenize every text object into words (then it's really easy to assemble lines below without having to break text, add its `options`, etc.)
	inputLines1.forEach(line => {
		line.forEach(cell => {
			let lineCells: TableCell[] = []
			let cellTextStr = cell.text + '' // force convert to string (compiled JS is better with this than a cast)
			let lineWords = cellTextStr.split(' ')
			let filteredWords = lineWords.filter(word => word)

			filteredWords.forEach((word, idx) => {
				let cellProps = { ...cell.options }
				// IMPORTANT: Handle `breakLine` prop - we cannot apply to each word - only apply to very last word!
				if (cellProps && cellProps.breakLine) cellProps.breakLine = idx + 1 === lineWords.length
				if (word) lineCells.push({ _type: SLIDE_OBJECT_TYPES.tablecell, text: word + (idx + 1 < filteredWords.length ? ' ' : ''), options: cellProps })
			})

			inputLines2.push(lineCells)
		})
	})
	if (verbose) {
		console.log(`[3/4] inputLines2 (${inputLines2.length})`)
		inputLines2.forEach(line => console.log(`line: ${JSON.stringify(line)}`))
		console.log('...............................................\n\n')
	}

	// STEP 4: Group cells/words into lines based upon space consumed by word letters
	inputLines2.forEach(line => {
		let lineCells: TableCell[] = []
		let strCurrLine = ''

		line.forEach(word => {
			// A: create new line when horizontal space is exhausted
			if (strCurrLine.length + word.text.length > CPL) {
				//if (verbose) console.log(`STEP 4: New line added: (${strCurrLine.length} + ${word.text.length} > ${CPL})`);
				parsedLines.push(lineCells)
				lineCells = [word]
				strCurrLine = word.text.toString()
			}

			// B: add current word to line cells
			lineCells.push(word)

			// C: add current word to `strCurrLine` which we use to keep track of line's char length
			strCurrLine += word.text.toString()
		})

		// Flush buffer: Only create a line when there's text to avoid empty row
		if (lineCells.length > 0) parsedLines.push(lineCells)
	})
	if (verbose) {
		console.log(`[4/4] parsedLines (${parsedLines.length})`)
		parsedLines.forEach((line, idx) => console.log(`[${idx + 1}] line: ${JSON.stringify(line)}`))
		console.log('...............................................\n\n')
	}

	// Done:
	return parsedLines
}

/**
 * Takes an array of table rows and breaks into an array of slides, which contain the calculated amount of table rows that fit on that slide
 * @param {TableCell[][]} tableRows - HTMLElementID of the table
 * @param {ITableToSlidesOpts} tabOpts - array of options (e.g.: tabsize)
 * @param {PresLayout} presLayout - Presentation layout
 * @param {SlideLayout} masterSlide - master slide (if any)
 * @return {TableRowSlide[]} array of table rows
 */
export function getSlidesForTableRows(tableRows: TableCell[][] = [], tabOpts: TableToSlidesProps = {}, presLayout: PresLayout, masterSlide?: SlideLayout): TableRowSlide[] {
	let arrInchMargins = DEF_SLIDE_MARGIN_IN,
		emuTabCurrH = 0,
		emuSlideTabW = EMU * 1,
		emuSlideTabH = EMU * 1,
		numCols = 0
	let tableRowSlides: TableRowSlide[] = []

	if (tabOpts.verbose) {
		console.log('[[VERBOSE MODE]]')
		console.log('|-- TABLE PROPS -----------------------------------------------------|')
		console.log(`| presLayout.height ......... = ${presLayout.height / EMU}`)
		console.log(`| tabOpts.h ................. = ${tabOpts.h}`)
		console.log(`| tabOpts.w ................. = ${tabOpts.w}`)
		console.log(`| tabOpts.colW .............. = ${tabOpts.colW}`)
		console.log(`| tabOpts.slideMargin ....... = ${tabOpts.slideMargin || ''}`)
		console.log('|-- CALCULATIONS ----------------------------------------------------|')
	}

	// STEP 1: Calculate margins
	{
		// Important: Use default size as zero cell margin is causing our tables to be too large and touch bottom of slide!
		if (!tabOpts.slideMargin && tabOpts.slideMargin !== 0) tabOpts.slideMargin = DEF_SLIDE_MARGIN_IN[0]

		if (masterSlide && typeof masterSlide._margin !== 'undefined') {
			if (Array.isArray(masterSlide._margin)) arrInchMargins = masterSlide._margin
			else if (!isNaN(Number(masterSlide._margin)))
				arrInchMargins = [Number(masterSlide._margin), Number(masterSlide._margin), Number(masterSlide._margin), Number(masterSlide._margin)]
		} else if (tabOpts.slideMargin || tabOpts.slideMargin === 0) {
			if (Array.isArray(tabOpts.slideMargin)) arrInchMargins = tabOpts.slideMargin
			else if (!isNaN(tabOpts.slideMargin)) arrInchMargins = [tabOpts.slideMargin, tabOpts.slideMargin, tabOpts.slideMargin, tabOpts.slideMargin]
		}

		if (tabOpts.verbose) console.log('| arrInchMargins ............ = ' + arrInchMargins.toString())
	}

	// STEP 2: Calculate number of columns
	{
		// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
		// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
		let firstRow = tableRows[0] || []
		firstRow.forEach(cell => {
			if (!cell) cell = { _type: SLIDE_OBJECT_TYPES.tablecell }
			let cellOpts = cell.options || null
			numCols += Number(cellOpts && cellOpts.colspan ? cellOpts.colspan : 1)
		})

		if (tabOpts.verbose) console.log('| numCols ................... = ' + numCols)
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

		if (tabOpts.verbose) console.log('| emuSlideTabW (in) ......... = ' + (emuSlideTabW / EMU).toFixed(1))
	}

	// STEP 5: Calculate column widths if not provided (emuSlideTabW will be used below to determine lines-per-col)
	if (!tabOpts.colW || !Array.isArray(tabOpts.colW)) {
		if (tabOpts.colW && !isNaN(Number(tabOpts.colW))) {
			let arrColW = []
			let firstRow = tableRows[0] || []
			firstRow.forEach(() => arrColW.push(tabOpts.colW))
			tabOpts.colW = []
			arrColW.forEach(val => {
				if (Array.isArray(tabOpts.colW)) tabOpts.colW.push(val)
			})
		}
		// No column widths provided? Then distribute cols.
		else {
			tabOpts.colW = []
			for (let iCol = 0; iCol < numCols; iCol++) {
				tabOpts.colW.push(emuSlideTabW / EMU / numCols)
			}
		}
	}

	// STEP 6: **MAIN** Iterate over rows, add table content, create new slides as rows overflow
	let newTableRowSlide: TableRowSlide = { rows: [] as TableRow[] }
	tableRows.forEach((row, iRow) => {
		// A: Row variables
		let rowCellLines: TableCell[] = []
		let maxCellMarTopEmu = 0
		let maxCellMarBtmEmu = 0
		let maxLineHeightEmu = 0

		// B: Create new row in data model
		let currTableRow: TableRow = []
		row.forEach(cell => {
			currTableRow.push({
				_type: SLIDE_OBJECT_TYPES.tablecell,
				text: [],
				options: cell.options,
			})

			if (cell.options.margin && cell.options.margin[0] && valToPts(cell.options.margin[0]) > maxCellMarTopEmu) maxCellMarTopEmu = valToPts(cell.options.margin[0])
			else if (tabOpts.margin && tabOpts.margin[0] && valToPts(tabOpts.margin[0]) > maxCellMarTopEmu) maxCellMarTopEmu = valToPts(tabOpts.margin[0])
			if (cell.options.margin && cell.options.margin[2] && valToPts(cell.options.margin[2]) > maxCellMarBtmEmu) maxCellMarBtmEmu = valToPts(cell.options.margin[2])
			else if (tabOpts.margin && tabOpts.margin[2] && valToPts(tabOpts.margin[2]) > maxCellMarBtmEmu) maxCellMarBtmEmu = valToPts(tabOpts.margin[2])
		})

		// C: Calc usable vertical space/table height. Set default value first, adjust below when necessary.
		emuSlideTabH =
			tabOpts.h && typeof tabOpts.h === 'number'
				? tabOpts.h
				: presLayout.height - inch2Emu(arrInchMargins[0] + arrInchMargins[2]) - (tabOpts.y && typeof tabOpts.y === 'number' ? tabOpts.y : 0)

		// D: RULE: Use margins for starting point after the initial Slide, not `opt.y` (ISSUE#43, ISSUE#47, ISSUE#48)
		if (tableRowSlides.length > 1 && typeof tabOpts.autoPageSlideStartY === 'number') {
			emuSlideTabH = tabOpts.h && typeof tabOpts.h === 'number' ? tabOpts.h : presLayout.height - inch2Emu(tabOpts.autoPageSlideStartY + arrInchMargins[2])
		} else if (tableRowSlides.length > 1 && typeof tabOpts.newSlideStartY === 'number') {
			// @deprecated v3.3.0
			emuSlideTabH = tabOpts.h && typeof tabOpts.h === 'number' ? tabOpts.h : presLayout.height - inch2Emu(tabOpts.newSlideStartY + arrInchMargins[2])
		} else if (tableRowSlides.length > 1 && typeof tabOpts.y === 'number') {
			emuSlideTabH = presLayout.height - inch2Emu((tabOpts.y / EMU < arrInchMargins[0] ? tabOpts.y / EMU : arrInchMargins[0]) + arrInchMargins[2])
			// Use whichever is greater: area between margins or the table H provided (dont shrink usable area - the whole point of over-riding X on paging is to *increarse* usable space)
			if (typeof tabOpts.h === 'number' && emuSlideTabH < tabOpts.h) emuSlideTabH = tabOpts.h
		} else if (typeof tabOpts.h === 'number' && typeof tabOpts.y === 'number')
			emuSlideTabH = tabOpts.h ? tabOpts.h : presLayout.height - inch2Emu((tabOpts.y / EMU || arrInchMargins[0]) + arrInchMargins[2])

		// E: --==[[ BUILD DATA SET ]]==-- (iterate over cells: split text into lines[], set `lineHeight`)
		row.forEach((cell, iCell) => {
			let newCell: TableCell = {
				_type: SLIDE_OBJECT_TYPES.tablecell,
				_lines: null,
				_lineHeight: inch2Emu(
					((cell.options && cell.options.fontSize ? cell.options.fontSize : tabOpts.fontSize ? tabOpts.fontSize : DEF_FONT_SIZE) *
						(LINEH_MODIFIER + (tabOpts.autoPageLineWeight ? tabOpts.autoPageLineWeight : 0))) /
						100
				),
				text: [],
				options: cell.options,
			}

			// 1: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
			if (newCell.options.rowspan) newCell._lineHeight = 0

			// 2: The parseTextToLines method uses `autoPageCharWeight`, so inherit from table options
			newCell.options.autoPageCharWeight = tabOpts.autoPageCharWeight ? tabOpts.autoPageCharWeight : null

			// 3: **MAIN** Parse cell contents into lines based upon col width, font, etc
			let totalColW = tabOpts.colW[iCell]
			if (cell.options.colspan && Array.isArray(tabOpts.colW)) {
				totalColW = tabOpts.colW.filter((_cell, idx) => idx >= iCell && idx < idx + cell.options.colspan).reduce((prev, curr) => prev + curr)
			}

			// 4: Create lines based upon available column width
			//newCell._lines = parseTextToLines(cell, totalColW / ONEPT, false)
			newCell._lines = parseTextToLines(cell, totalColW / ONEPT, true) // FIXME: HTML and Slide7Demo are CRAP

			// 5: Add cell to array
			rowCellLines.push(newCell)
		})

		// F: Start row height with margins
		emuTabCurrH += maxCellMarTopEmu + maxCellMarBtmEmu
		//if (tabOpts.verbose) console.log(`- SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: maxCellMarTopEmu=${maxCellMarTopEmu} / maxCellMarBtmEmu=${maxCellMarBtmEmu}`)

		/** G: --==[[ PAGE DATA SET ]]==--
		 * Add text one-line-a-time to this row's cells until: lines are exhausted OR table height limit is hit
		 *
		 * Design:
		 * - Building cells L-to-R/loop style wont work as one could be 100 lines and another 1 line
		 * - Therefore, build the whole row, one-line-at-a-time, across each table columns
		 * - Then, when the vertical size limit is hit is by any of the cells, make a new slide and continue adding any remaining lines
		 *
		 * Implementation:
		 * `rowCellLines` is an array of cells
		 * - each cell contains an array of lines
		 * EX:
		 * {
		 *    _lines: [{ text:'cell 1' }, { text:'line 2' }],
		 *    _lines: [{ text:'cell 2' }, { text:'line 2' }],
		 *    _lines: [{ text:'cell 3' }, { text:'line 2' }, { text:'line 3' }, { text:'line 4' }],
		 * }
		 */
		if (rowCellLines) {
			if (tabOpts.verbose) console.log(`\n| SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: START...`)

			rowCellLines.forEach((cell, cellIdx) => {
				cell._lines.forEach((line, lineIdx) => {
					// A: create a new slide if there is insufficient room for the current row
					if (emuTabCurrH + maxLineHeightEmu > emuSlideTabH) {
						if (tabOpts.verbose) {
							console.log('\n|--------------------------------------------------------------------|')
							console.log(
								`|-- NEW SLIDE CREATED (b/c: currH + newRowH > maxH) => ${(emuTabCurrH / EMU).toFixed(2)} + ${(maxLineHeightEmu / EMU).toFixed(2)} > ${
									emuSlideTabH / EMU
								}`
							)
							console.log('|--------------------------------------------------------------------|\n\n')
						}

						// 1: add current row slide or it will be lost
						newTableRowSlide.rows.push(currTableRow)

						// 2: add current slide to Slides array
						tableRowSlides.push(newTableRowSlide)

						// 3: reset working/curr slide to hold rows as they're created
						let newRows: TableRow[] = []
						newTableRowSlide = { rows: newRows }

						// 4: reset working/curr row
						currTableRow = []
						row.forEach(cell => {
							currTableRow.push({
								_type: SLIDE_OBJECT_TYPES.tablecell,
								text: [],
								options: cell.options,
							})
						})

						// 5: reset current table height for this new Slide
						emuTabCurrH = 0

						// FIXME: this isnt working on HTML-slides demo!!! 202010821
						// 6: handle repeat headers option /or/ Add new empty row to continue current lines into
						if ((tabOpts.addHeaderToEach || tabOpts.autoPageRepeatHeader) && tabOpts._arrObjTabHeadRows) {
							let tableHeadRows: TableCell[][] = []
							tabOpts._arrObjTabHeadRows.forEach(row => {
								let newHeadRow = []
								row.forEach(cell => newHeadRow.push(cell))
								tableHeadRows.push(newHeadRow)
							})
							tableRows = [...tableHeadRows, ...tableRows]
						}
					}

					// B: get current cell on `currTableRow`
					let currCell = currTableRow[cellIdx]

					// C: create new line (add all words)
					if (Array.isArray(currCell.text)) currCell.text = currCell.text.concat(line)

					// D: increase current table row height
					if (cell._lineHeight > maxLineHeightEmu) maxLineHeightEmu = cell._lineHeight

					// E: increase table height by the max height from above
					emuTabCurrH += maxLineHeightEmu

					// DONE
					if (tabOpts.verbose) {
						console.log(
							`- SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: CELL [${cellIdx}]: LINE [${lineIdx}] added ... emuTabCurrH = ${(emuTabCurrH / EMU).toFixed(2)}`
						)
					}
				})
			})
		}

		// 7: Flush/capture row buffer before it resets at the top of this loop
		if (currTableRow.length > 0) newTableRowSlide.rows.push(currTableRow)

		if (tabOpts.verbose)
			console.log(
				`- SLIDE [${tableRowSlides.length}]: ROW [${iRow}]: ...COMPLETE ...... emuTabCurrH = ${(emuTabCurrH / EMU).toFixed(2)} ( emuSlideTabH = ${(
					emuSlideTabH / EMU
				).toFixed(2)} )`
			)
	})

	// STEP 7: Flush buffer / add final slide
	tableRowSlides.push(newTableRowSlide)

	if (tabOpts.verbose) {
		console.log(`\n|================================================|\n| FINAL: tableRowSlides.length = ${tableRowSlides.length}`)
		console.log(tableRowSlides)
		//console.log(JSON.stringify(tableRowSlides,null,2))
		console.log(`|================================================|\n\n`)
	}

	// LAST:
	return tableRowSlides
}

/**
 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
 * @param {PptxGenJS} pptx - pptxgenjs instance
 * @param {string} tabEleId - HTMLElementID of the table
 * @param {ITableToSlidesOpts} options - array of options (e.g.: tabsize)
 * @param {SlideLayout} masterSlide - masterSlide
 */
export function genTableToSlides(pptx: PptxGenJS, tabEleId: string, options: TableToSlidesProps = {}, masterSlide?: SlideLayout) {
	let opts = options || {}
	opts.slideMargin = opts.slideMargin || opts.slideMargin === 0 ? opts.slideMargin : 0.5
	let emuSlideTabW = opts.w || pptx.presLayout.width
	let arrObjTabHeadRows: [TableCell[]?] = []
	let arrObjTabBodyRows: [TableCell[]?] = []
	let arrObjTabFootRows: [TableCell[]?] = []
	let arrColW: number[] = []
	let arrTabColW: number[] = []
	let arrInchMargins: [number, number, number, number] = [0.5, 0.5, 0.5, 0.5] // TRBL-style
	let intTabW = 0

	// REALITY-CHECK:
	if (!document.getElementById(tabEleId)) throw new Error('tableToSlides: Table ID "' + tabEleId + '" does not exist!')

	// STEP 1: Set margins
	if (masterSlide && masterSlide._margin) {
		if (Array.isArray(masterSlide._margin)) arrInchMargins = masterSlide._margin
		else if (!isNaN(masterSlide._margin)) arrInchMargins = [masterSlide._margin, masterSlide._margin, masterSlide._margin, masterSlide._margin]
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

	// STEP 2: Grab table col widths - just find the first availble row, either thead/tbody/tfoot, others may have colspsna,s who cares, we only need col widths from 1
	let firstRowCells = document.querySelectorAll(`#${tabEleId} tr:first-child th`)
	if (firstRowCells.length === 0) firstRowCells = document.querySelectorAll(`#${tabEleId} tr:first-child td`)
	firstRowCells.forEach((cell: HTMLElement) => {
		if (cell.getAttribute('colspan')) {
			// Guesstimate (divide evenly) col widths
			// NOTE: both j$query and vanilla selectors return {0} when table is not visible)
			for (let idxc = 0; idxc < Number(cell.getAttribute('colspan')); idxc++) {
				arrTabColW.push(Math.round(cell.offsetWidth / Number(cell.getAttribute('colspan'))))
			}
		} else {
			arrTabColW.push(cell.offsetWidth)
		}
	})
	arrTabColW.forEach(colW => {
		intTabW += colW
	})

	// STEP 3: Calc/Set column widths by using same column width percent from HTML table
	arrTabColW.forEach((colW, idxW) => {
		let intCalcWidth = Number(((Number(emuSlideTabW) * ((colW / intTabW) * 100)) / 100 / EMU).toFixed(2))
		let intMinWidth = 0
		let colSelectorMin = document.querySelector(`#${tabEleId} thead tr:first-child th:nth-child(${idxW + 1})`)
		if (colSelectorMin) intMinWidth = Number(colSelectorMin.getAttribute('data-pptx-min-width'))
		let intSetWidth = 0
		let colSelectorSet = document.querySelector(`#${tabEleId} thead tr:first-child th:nth-child(${idxW + 1})`)
		if (colSelectorSet) intMinWidth = Number(colSelectorSet.getAttribute('data-pptx-width'))
		arrColW.push(intSetWidth ? intSetWidth : intMinWidth > intCalcWidth ? intMinWidth : intCalcWidth)
	})
	if (opts.verbose) {
		console.log(`arrColW ................ = ${arrColW.toString()}`)
	}

	// STEP 4: Iterate over each table element and create data arrays (text and opts)
	// NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
	;['thead', 'tbody', 'tfoot'].forEach(part => {
		document.querySelectorAll(`#${tabEleId} ${part} tr`).forEach((row: HTMLTableRowElement) => {
			let arrObjTabCells: TableCell[] = []
			Array.from(row.cells).forEach(cell => {
				// A: Get RGB text/bkgd colors
				let arrRGB1 = window.getComputedStyle(cell).getPropertyValue('color').replace(/\s+/gi, '').replace('rgba(', '').replace('rgb(', '').replace(')', '').split(',')
				let arrRGB2 = window
					.getComputedStyle(cell)
					.getPropertyValue('background-color')
					.replace(/\s+/gi, '')
					.replace('rgba(', '')
					.replace('rgb(', '')
					.replace(')', '')
					.split(',')
				if (
					// NOTE: (ISSUE#57): Default for unstyled tables is black bkgd, so use white instead
					window.getComputedStyle(cell).getPropertyValue('background-color') === 'rgba(0, 0, 0, 0)' ||
					window.getComputedStyle(cell).getPropertyValue('transparent')
				) {
					arrRGB2 = ['255', '255', '255']
				}

				// B: Create option object
				let cellOpts: TableCellProps = {
					align: null,
					bold:
						window.getComputedStyle(cell).getPropertyValue('font-weight') === 'bold' ||
						Number(window.getComputedStyle(cell).getPropertyValue('font-weight')) >= 500
							? true
							: false,
					border: null,
					color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
					fill: { color: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2])) },
					fontFace:
						(window.getComputedStyle(cell).getPropertyValue('font-family') || '').split(',')[0].replace(/"/g, '').replace('inherit', '').replace('initial', '') ||
						null,
					fontSize: Number(window.getComputedStyle(cell).getPropertyValue('font-size').replace(/[a-z]/gi, '')),
					margin: null,
					colspan: Number(cell.getAttribute('colspan')) || null,
					rowspan: Number(cell.getAttribute('rowspan')) || null,
					valign: null,
				}

				if (['left', 'center', 'right', 'start', 'end'].indexOf(window.getComputedStyle(cell).getPropertyValue('text-align')) > -1) {
					let align = window.getComputedStyle(cell).getPropertyValue('text-align').replace('start', 'left').replace('end', 'right')
					cellOpts.align = align === 'center' ? 'center' : align === 'left' ? 'left' : align === 'right' ? 'right' : null
				}
				if (['top', 'middle', 'bottom'].indexOf(window.getComputedStyle(cell).getPropertyValue('vertical-align')) > -1) {
					let valign = window.getComputedStyle(cell).getPropertyValue('vertical-align')
					cellOpts.valign = valign === 'top' ? 'top' : valign === 'middle' ? 'middle' : valign === 'bottom' ? 'bottom' : null
				}

				// C: Add padding [margin] (if any)
				// NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
				if (window.getComputedStyle(cell).getPropertyValue('padding-left')) {
					cellOpts.margin = [0, 0, 0, 0]
					let sidesPad = ['padding-top', 'padding-right', 'padding-bottom', 'padding-left']
					sidesPad.forEach((val, idxs) => {
						cellOpts.margin[idxs] = Math.round(Number(window.getComputedStyle(cell).getPropertyValue(val).replace(/\D/gi, '')))
					})
				}

				// D: Add border (if any)
				if (
					window.getComputedStyle(cell).getPropertyValue('border-top-width') ||
					window.getComputedStyle(cell).getPropertyValue('border-right-width') ||
					window.getComputedStyle(cell).getPropertyValue('border-bottom-width') ||
					window.getComputedStyle(cell).getPropertyValue('border-left-width')
				) {
					cellOpts.border = [null, null, null, null]
					let sidesBor = ['top', 'right', 'bottom', 'left']
					sidesBor.forEach((val, idxb) => {
						let intBorderW = Math.round(
							Number(
								window
									.getComputedStyle(cell)
									.getPropertyValue('border-' + val + '-width')
									.replace('px', '')
							)
						)
						let arrRGB = []
						arrRGB = window
							.getComputedStyle(cell)
							.getPropertyValue('border-' + val + '-color')
							.replace(/\s+/gi, '')
							.replace('rgba(', '')
							.replace('rgb(', '')
							.replace(')', '')
							.split(',')
						let strBorderC = rgbToHex(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]))
						cellOpts.border[idxb] = { pt: intBorderW, color: strBorderC }
					})
				}

				// LAST: Add cell
				arrObjTabCells.push({
					_type: SLIDE_OBJECT_TYPES.tablecell,
					text: cell.innerText, // `innerText` returns <br> as "\n", so linebreak etc. work later!
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
					console.log(`table parsing: unexpected table part: ${part}`)
					break
			}
		})
	})

	// STEP 5: Break table into Slides as needed
	// Pass head-rows as there is an option to add to each table and the parse func needs this data to fulfill that option
	opts._arrObjTabHeadRows = arrObjTabHeadRows || null
	opts.colW = arrColW
	getSlidesForTableRows([...arrObjTabHeadRows, ...arrObjTabBodyRows, ...arrObjTabFootRows], opts, pptx.presLayout, masterSlide).forEach((slide, idxTr) => {
		// A: Create new Slide
		let newSlide = pptx.addSlide({ masterName: opts.masterSlideName || null })

		// B: DESIGN: Reset `y` to startY or margin after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
		if (idxTr === 0) opts.y = opts.y || arrInchMargins[0]
		if (idxTr > 0) opts.y = opts.autoPageSlideStartY || opts.newSlideStartY || arrInchMargins[0]
		if (opts.verbose) console.log('opts.autoPageSlideStartY:' + opts.autoPageSlideStartY + ' / arrInchMargins[0]:' + arrInchMargins[0] + ' => opts.y = ' + opts.y)

		// C: Add table to Slide
		newSlide.addTable(slide.rows, { x: opts.x || arrInchMargins[3], y: opts.y, w: Number(emuSlideTabW) / EMU, colW: arrColW, autoPage: false })

		// D: Add any additional objects
		if (opts.addImage) newSlide.addImage({ path: opts.addImage.url, x: opts.addImage.x, y: opts.addImage.y, w: opts.addImage.w, h: opts.addImage.h })
		if (opts.addShape) newSlide.addShape(opts.addShape.shape, opts.addShape.options || {})
		if (opts.addTable) newSlide.addTable(opts.addTable.rows, opts.addTable.options || {})
		if (opts.addText) newSlide.addText(opts.addText.text, opts.addText.options || {})
	})
}
