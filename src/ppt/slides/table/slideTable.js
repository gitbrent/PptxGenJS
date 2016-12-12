import { inch2Emu, rgbToHex, parseTextToLines } from '../utils/helpers'
import Slide from '../slide.js'

const ONEPT= 12700, EMU=914400;

export default class SlideTable {

    addSlidesForTable( tabEleId, inOpts ) {

        let opts = inOpts || {},
            arrObjTabHeadRows = [],
            arrObjTabBodyRows = [],
            arrObjTabFootRows = [],
            arrObjSlides = [],
            arrRows = [],
            arrColW = [],
            arrTabColW = [],
            intTabW = 0,
            emuTabCurrH = 0;

        // NOTE: Look for opts.margin first as user can override Slide Master settings if they want
        let arrInchMargins = this.lookMargin(opts);

        let emuSlideTabW = ( Slide.gObjPptx.pptLayout.width - inch2Emu( arrInchMargins[ 1 ] + arrInchMargins[ 3 ] ) );

        let emuSlideTabH = ( Slide.gObjPptx.pptLayout.height - inch2Emu( arrInchMargins[ 0 ] + arrInchMargins[ 2 ] ) );

        // STEP 1: Grab overall table style/col widths
        this.tableStyle(tabEleId, arrTabColW);
        $.each( arrTabColW, function( i, colW ) {
            intTabW += colW;
        } );

        // STEP 2: Calc/Set column widths by using same column width percent from HTML table
        this.calcColWidth(arrTabColW, tabEleId, arrColW, emuSlideTabW, intTabW);

        // STEP 3: Iterate over each table element and create data arrays (text and opts)
        // NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
        this.createDataArray(tabEleId, arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows);

        // STEP 4: Paginate data: Iterate over all table rows, divide into slides/pages based upon the row height>overall height
        this.paginateData(arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows, arrColW, emuTabCurrH, emuSlideTabH, arrRows, arrObjSlides, opts); // tab loop
        // Flush final row buffer to slide
        arrObjSlides.push( $.extend( true, [], arrRows ) );

        // STEP 5: Create a SLIDE for each of our 1-N table pieces
        this.createSlides(arrObjSlides, opts, arrInchMargins, emuSlideTabW, arrColW);
    }

    tableStyle(tabEleId, arrTabColW) {
        $.each(['thead', 'tbody', 'tfoot'], (i, val) => {
            if ($('#' + tabEleId + ' ' + val + ' tr').length > 0) {
                $('#' + tabEleId + ' ' + val + ' tr:first-child').find('th, td').each(function (i, cell) {
                    // TODO 1.5: This is a hack - guessing at col widths when colspan
                    if ($(this).attr('colspan')) {
                        for (var idx = 0; idx < $(this).attr('colspan'); idx++) {
                            arrTabColW.push(Math.round($(this).outerWidth() / $(this).attr('colspan')));
                        }
                    } else {
                        arrTabColW.push($(this).outerWidth());
                    }
                });
                return false; // break out of .each loop
            }
        });
    }
    calcColWidth(arrTabColW, tabEleId, arrColW, emuSlideTabW, intTabW) {
        $.each(arrTabColW, function (i, colW) {
            ( $('#' + tabEleId + ' thead tr:first-child th:nth-child(' + ( i + 1 ) + ')').data('pptx-min-width') ) ?
                arrColW.push(inch2Emu($('#' + tabEleId + ' thead tr:first-child th:nth-child(' + ( i + 1 ) + ')').data('pptx-min-width'))) : arrColW.push(Math.round(( emuSlideTabW * ( colW / intTabW * 100 ) ) / 100));
        });
    }
    createDataArray(tabEleId, arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows) {
        $.each(['thead', 'tbody', 'tfoot'], function (i, val) {
            $('#' + tabEleId + ' ' + val + ' tr').each(function (i, row) {
                var arrObjTabCells = [];
                $(row).find('th, td').each(function (i, cell) {
                    // A: Covert colors to Hex from RGB
                    var arrRGB1 = [];
                    var arrRGB2 = [];
                    arrRGB1 = $(cell).css('color').replace(/\s+/gi, '').replace('rgb(', '').replace(')', '').split(',');
                    arrRGB2 = $(cell).css('background-color').replace(/\s+/gi, '').replace('rgb(', '').replace(')', '').split(',');

                    // B: Create option object
                    var objOpts = {
                        font_size: $(cell).css('font-size').replace(/\D/gi, ''),
                        bold: ( ( $(cell).css('font-weight') == "bold" || Number($(cell).css('font-weight')) >= 500 ) ? true : false ),
                        color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
                        fill: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2]))
                    };
                    if ($.inArray($(cell).css('text-align'), ['left', 'center', 'right', 'start', 'end']) > -1) objOpts.align = $(cell).css('text-align').replace('start', 'left').replace('end', 'right');
                    if ($.inArray($(cell).css('vertical-align'), ['top', 'middle', 'bottom']) > -1) objOpts.valign = $(cell).css('vertical-align');

                    // C: Add padding [margin] (if any)
                    // NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
                    if ($(cell).css('padding-left')) {
                        objOpts.marginPt = [];
                        $.each(['padding-top', 'padding-right', 'padding-bottom', 'padding-left'], function (i, val) {
                            objOpts.marginPt.push(Math.round($(cell).css(val).replace(/\D/gi, '') * ONEPT));
                        });
                    }

                    // D: Add colspan (if any)
                    if ($(cell).attr('colspan')) objOpts.colspan = $(cell).attr('colspan');

                    // E: Add border (if any)
                    if ($(cell).css('border-top-width') || $(cell).css('border-right-width') || $(cell).css('border-bottom-width') || $(cell).css('border-left-width')) {
                        objOpts.border = [];
                        $.each(['top', 'right', 'bottom', 'left'], function (i, val) {
                            var intBorderW = Math.round(Number($(cell).css('border-' + val + '-width').replace('px', '')));
                            var arrRGB = [];
                            arrRGB = $(cell).css('border-' + val + '-color').replace(/\s+/gi, '').replace('rgba(', '').replace('rgb(', '').replace(')', '').split(',');
                            var strBorderC = rgbToHex(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]));
                            objOpts.border.push({
                                pt: intBorderW,
                                color: strBorderC
                            });
                        });
                    }

                    // F: Massage cell text so we honor linebreak tag as a line break during line parsing
                    var $cell = $(cell).clone();
                    $cell.html($(cell).html().replace(/<br[^>]*>/gi, '\n'));

                    // LAST: Add cell
                    arrObjTabCells.push({
                        text: $cell.text(),
                        opts: objOpts
                    });
                });
                switch (val) {
                    case 'thead':
                        arrObjTabHeadRows.push(arrObjTabCells);
                        break;
                    case 'tbody':
                        arrObjTabBodyRows.push(arrObjTabCells);
                        break;
                    case 'tfoot':
                        arrObjTabFootRows.push(arrObjTabCells);
                        break;
                    default:
                }
            });
        });
    }
    paginateData(arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows, arrColW, emuTabCurrH, emuSlideTabH, arrRows, arrObjSlides, opts) {
        $.each([arrObjTabHeadRows, arrObjTabBodyRows, arrObjTabFootRows], function (iTab, tab) {
            var currRow = [];
            $.each(tab, function (iRow, row) {
                // A: Reset ROW variables
                var arrCellsLines = [],
                    arrCellsLineHeights = [],
                    emuRowH = 0,
                    intMaxLineCnt = 0,
                    intMaxColIdx = 0;

                // B: Parse and store each cell's text into line array (*MAGIC HAPPENS HERE*)
                $(row).each(function (iCell, cell) {
                    // 1: Create a cell object for each table column
                    currRow.push({
                        text: '',
                        opts: cell.opts
                    });

                    // 2: Parse cell contents into lines (**MAGIC HAPENSS HERE**)
                    var lines = parseTextToLines(cell.text, cell.opts.font_size, ( arrColW[iCell] / ONEPT ));
                    arrCellsLines.push(lines);

                    // 3: Keep track of max line count within all row cells
                    if (lines.length > intMaxLineCnt) {
                        intMaxLineCnt = lines.length;
                        intMaxColIdx = iCell;
                    }
                });

                // C: Calculate Line-Height
                // FYI: Line-Height =~ font-size [px~=pt] * 1.65 / 100 = inches high
                // FYI: 1px = 14288 EMU (0.156 inches) @96 PPI - I ended up going with 20000 EMU as margin spacing needed a bit more than 1:1
                $(row).each(function (iCell, cell) {
                    var lineHeight = inch2Emu(cell.opts.font_size * 1.65 / 100);
                    if (Array.isArray(cell.opts.marginPt) && cell.opts.marginPt[0]) lineHeight += cell.opts.marginPt[0] / intMaxLineCnt;
                    if (Array.isArray(cell.opts.marginPt) && cell.opts.marginPt[2]) lineHeight += cell.opts.marginPt[2] / intMaxLineCnt;
                    arrCellsLineHeights.push(Math.round(lineHeight));
                });

                // D: AUTO-PAGING: Add text one-line-a-time to this row's cells until: lines are exhausted OR table H limit is hit
                for (var idx = 0; idx < intMaxLineCnt; idx++) {
                    // 1: Add the current line to cell
                    for (var col = 0; col < arrCellsLines.length; col++) {
                        // A: Commit this slide to Presenation if table Height limit is hit
                        if (emuTabCurrH + arrCellsLineHeights[intMaxColIdx] > emuSlideTabH) {
                            // 1: Add the current row to table
                            // NOTE: Edge cases can occur where we create a new slide only to have no more lines
                            // ....: and then a blank row sits at the bottom of a table!
                            // ....: Hence, we very all cells have text before adding this final row.
                            $.each(currRow, function (i, cell) {
                                if (cell.text.length > 0) {
                                    // IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
                                    arrRows.push($.extend(true, [], currRow));
                                    return false; // break out of .each loop
                                }
                            });
                            // 2: Add new Slide with current array of table rows
                            arrObjSlides.push($.extend(true, [], arrRows));
                            // 3: Empty rows for new Slide
                            arrRows.length = 0;
                            // 4: Reset curr table height for new Slide
                            emuTabCurrH = 0; // This row's emuRowH w/b added below
                            // 5: Empty current row's text (continue adding lines where we left off below)
                            $.each(currRow, function (i, cell) {
                                cell.text = '';
                            });
                            // 6: Auto-Paging Options: addHeaderToEach
                            if (opts.addHeaderToEach) {
                                var headRow = [];
                                $.each(arrObjTabHeadRows[0], function (iCell, cell) {
                                    headRow.push({
                                        text: cell.text,
                                        opts: cell.opts
                                    });
                                    var lines = parseTextToLines(cell.text, cell.opts.font_size, ( arrColW[iCell] / ONEPT ));
                                    if (lines.length > intMaxLineCnt) {
                                        intMaxLineCnt = lines.length;
                                        intMaxColIdx = iCell;
                                    }
                                });
                                arrRows.push($.extend(true, [], headRow));
                            }
                        }

                        // B: Add next line of text to this cell
                        if (arrCellsLines[col][idx]) currRow[col].text += arrCellsLines[col][idx];
                    }

                    // 2: Add this new rows H to overall (The cell with the longest line array is the one we use as the determiner for overall row Height)
                    emuTabCurrH += arrCellsLineHeights[intMaxColIdx];
                }

                // E: Flush row buffer - Add the current row to table, then truncate row cell array
                // IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
                arrRows.push($.extend(true, [], currRow));
                currRow.length = 0;
            }); // row loop
        });
    }
    createSlides(arrObjSlides, opts, arrInchMargins, emuSlideTabW, arrColW) {
        $.each(arrObjSlides, (i, slide) => {
            // A: Create table row array
            let arrTabRows = [];

            // B: Create new Slide
            let newSlide = ( opts && opts.master && gObjPptxMasters ) ? new Slide().addNewSlide(opts.master) : new Slide().addNewSlide();

            // C: Create array of Rows
            $.each(slide, (i, row) => {
                var arrTabRowCells = [];
                $.each(row, (i, cell) => {
                    arrTabRowCells.push(cell);
                });
                arrTabRows.push(arrTabRowCells);
            });

            // D: Add table to Slide
            newSlide.addTable(arrTabRows, {
                x: arrInchMargins[3],
                y: arrInchMargins[0],
                cx: ( emuSlideTabW / EMU )
            }, {
                colW: arrColW
            });

            // E: Add any additional objects
            if (opts.addImage) newSlide.addImage(opts.addImage.url, opts.addImage.x, opts.addImage.y, opts.addImage.w, opts.addImage.h);
            if (opts.addText) newSlide.addText(opts.addText.text, ( opts.addText.opts || {} ));
            if (opts.addShape) newSlide.addShape(opts.addShape.shape, ( opts.addShape.opts || {} ));
            if (opts.addTable) newSlide.addTable(opts.addTable.rows, ( opts.addTable.opts || {} ), ( opts.addTable.tabOpts || {} ));
        });
    }
    lookMargin(opts) {
        var arrInchMargins = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
        if (opts && opts.margin) {
            if (Array.isArray(opts.margin)) arrInchMargins = opts.margin;
            else if (!isNaN(opts.margin)) arrInchMargins = [opts.margin, opts.margin, opts.margin, opts.margin];
        } else if (opts && opts.master && opts.master.margin && gObjPptxMasters) {
            if (Array.isArray(opts.master.margin)) arrInchMargins = opts.master.margin;
            else if (!isNaN(opts.master.margin)) arrInchMargins = [opts.master.margin, opts.master.margin, opts.master.margin, opts.master.margin];
        }
        return arrInchMargins;
    }

}




