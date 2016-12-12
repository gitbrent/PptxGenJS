import { inch2Emu, decodeXmlEntities, parseTextToLines } from '../helpers'


export default function ExportTable(inSlide, slideObj, intTableNum, x, y, cx, cy){

    const ONEPT = 12700, EMU = 914400;
    let arrRowspanCells = [],
        arrTabRows = slideObj.arrTabRows,
        objTabOpts = slideObj.objTabOpts,
        intColCnt = 0,
        intColW = 0;

    // NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
    // ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
    for (var tmp=0; tmp<arrTabRows[0].length; tmp++) {
        intColCnt += ( arrTabRows[0][tmp].opts && arrTabRows[0][tmp].opts.colspan ) ? Number(arrTabRows[0][tmp].opts.colspan) : 1;
    }

    // STEP 1: Start Table XML
    // NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
    var strXml = '<p:graphicFrame>'
        + '  <p:nvGraphicFramePr>'
        + '    <p:cNvPr id="'+ (intTableNum*inSlide.numb + 1) +'" name="Table '+ (intTableNum*inSlide.numb) +'"/>'
        + '    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
        + '    <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>'
        + '  </p:nvGraphicFramePr>'
        + '  <p:xfrm>'
        + '    <a:off  x="'+ (x  || EMU) +'"  y="'+ (y  || EMU) +'"/>'
        + '    <a:ext cx="'+ (cx || EMU) +'" cy="'+ (cy || EMU) +'"/>'
        + '  </p:xfrm>'
        + '  <a:graphic>'
        + '    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        + '      <a:tbl>'
        + '        <a:tblPr/>';
    // + '        <a:tblPr bandRow="1"/>';
    // TODO 1.5: Support banded rows, first/last row, etc.
    // NOTE: Banding, etc. only shows when using a table style! (or set alt row color if banding)
    // <a:tblPr firstCol="0" firstRow="0" lastCol="0" lastRow="0" bandCol="0" bandRow="1">

    // STEP 2: Set column widths
    // Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
    // A: Col widths provided?
    if ( Array.isArray(objTabOpts.colW) ) {
        strXml += '<a:tblGrid>';
        for ( var col=0; col<intColCnt; col++ ) {
            strXml += '  <a:gridCol w="'+ (objTabOpts.colW[col] || (slideObj.options.cx/intColCnt)) +'"/>';
        }
        strXml += '</a:tblGrid>';
    }
    // B: Table Width provided without colW? Then distribute cols
    else {
        intColW = (objTabOpts.colW) ? objTabOpts.colW : EMU;
        if ( slideObj.options.cx && !objTabOpts.colW ) intColW = ( slideObj.options.cx / intColCnt );
        strXml += '<a:tblGrid>';
        for ( var col=0; col<intColCnt; col++ ) { strXml += '<a:gridCol w="'+ intColW +'"/>'; }
        strXml += '</a:tblGrid>';
    }
    // C: Table Height provided without rowH? Then distribute rows
    var intRowH = (objTabOpts.rowH) ? inch2Emu(objTabOpts.rowH) : 0;
    if ( slideObj.options.cy && !objTabOpts.rowH ) intRowH = ( slideObj.options.cy / arrTabRows.length );

    // STEP 3: Build an array of rowspan cells now so we can add stubs in as we loop below
    $.each(arrTabRows, (rIdx,row) => {
        $(row).each(function(cIdx,cell){
            var colIdx = cIdx;
            if ( cell.opts && cell.opts.rowspan && Number.isInteger(cell.opts.rowspan) ) {
                for (let idy=1; idy<cell.opts.rowspan; idy++) {
                    arrRowspanCells.push( {row:(rIdx+idy), col:colIdx} );
                    colIdx++; // For cases where we already have a rowspan in this row - we need to Increment to account for this extra cell!
                }
            }
        });
    });

    // STEP 4: Build table rows/cells
    $.each(arrTabRows, (rIdx,row) => {
        if ( Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx] ) intRowH = inch2Emu(objTabOpts.rowH[rIdx]);

        // A: Start row
        strXml += '<a:tr h="'+ intRowH +'">';

        // B: Loop over each CELL
        $(row).each(function(cIdx,cell){
            // 1: OPTIONS: Build/set cell options (blocked for code folding)
            {
                // 1: Load/Create options
                var cellOpts = cell.opts || {};

                // 2: Do Important/Override Opts
                // Feature: TabOpts Default Values (tabOpts being used when cellOpts dont exist):
                // SEE: http://officeopenxml.com/drwTableCellProperties-alignment.php
                $.each(['align','bold','border','color','fill','font_face','font_size','underline','valign'], function(i,name){
                    if ( objTabOpts[name] && ! cellOpts[name]) cellOpts[name] = objTabOpts[name];
                });

                var cellB       = (cellOpts.bold)       ? ' b="1"' : ''; // [0,1] or [false,true]
                var cellU       = (cellOpts.underline)  ? ' u="sng"' : ''; // [none,sng (single), et al.]
                var cellFont    = (cellOpts.font_face)  ? ' <a:latin typeface="'+ cellOpts.font_face +'"/>' : '';
                var cellFontPt  = (cellOpts.font_size)  ? ' sz="'+ cellOpts.font_size +'00"' : '';
                var cellAlign   = (cellOpts.align)      ? ' algn="'+ cellOpts.align.replace(/^c$/i,'ctr').replace('center','ctr').replace('left','l').replace('right','r') +'"' : '';
                var cellValign  = (cellOpts.valign)     ? ' anchor="'+ cellOpts.valign.replace(/^c$/i,'ctr').replace(/^m$/i,'ctr').replace('center','ctr').replace('middle','ctr').replace('top','t').replace('btm','b').replace('bottom','b') +'"' : '';
                var cellColspan = (cellOpts.colspan)    ? ' gridSpan="'+ cellOpts.colspan +'"' : '';
                var cellRowspan = (cellOpts.rowspan)    ? ' rowSpan="'+ cellOpts.rowspan +'"' : '';
                var cellFontClr = ((cell.optImp && cell.optImp.color) || cellOpts.color) ? ' <a:solidFill><a:srgbClr val="'+ ((cell.optImp && cell.optImp.color) || cellOpts.color) +'"/></a:solidFill>' : '';
                var cellFill    = ((cell.optImp && cell.optImp.fill)  || cellOpts.fill ) ? ' <a:solidFill><a:srgbClr val="'+ ((cell.optImp && cell.optImp.fill) || cellOpts.fill) +'"/></a:solidFill>' : '';
                var intMarginPt = (cellOpts.marginPt || cellOpts.marginPt == 0) ? (cellOpts.marginPt * ONEPT) : 0;
                // Margin/Padding:
                var cellMargin  = '';
                if ( cellOpts.marginPt && Array.isArray(cellOpts.marginPt) ) {
                    cellMargin = ' marL="'+ cellOpts.marginPt[3] +'" marR="'+ cellOpts.marginPt[1] +'" marT="'+ cellOpts.marginPt[0] +'" marB="'+ cellOpts.marginPt[2] +'"';
                }
                else if ( cellOpts.marginPt && Number.isInteger(cellOpts.marginPt) ) {
                    cellMargin = ' marL="'+ intMarginPt +'" marR="'+ intMarginPt +'" marT="'+ intMarginPt +'" marB="'+ intMarginPt +'"';
                }
            }

            // 2: Cell Content: Either the text element or the cell itself (for when users just pass a string - no object or options)
            var strCellText = ((typeof cell === 'object') ? cell.text : cell);

            // TODO 1.5: Cell NOWRAP property (text wrap: add to a:tcPr (horzOverflow="overflow" or whatev opts exist)

            // 3: ROWSPAN: Add dummy cells for any active rowspan
            // TODO 1.5: ROWSPAN & COLSPAN in same cell is not yet handled!
            if ( arrRowspanCells.filter(function(obj){ return obj.row == rIdx && obj.col == cIdx }).length > 0 ) {
                strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>';
            }

            // 4: Start Table Cell, add Align, add Text content
            strXml += ' <a:tc'+ cellColspan + cellRowspan +'>'
                + '  <a:txBody>'
                + '    <a:bodyPr/>'
                + '    <a:lstStyle/>'
                + '    <a:p>'
                + '      <a:pPr'+ cellAlign +'/>'
                + '      <a:r>'
                + '        <a:rPr lang="en-US" dirty="0" smtClean="0"'+ cellFontPt + cellB + cellU +'>'+ cellFontClr + cellFont +'</a:rPr>'
                + '        <a:t>'+ decodeXmlEntities(strCellText) +'</a:t>'
                + '      </a:r>'
                + '      <a:endParaRPr lang="en-US" dirty="0"/>'
                + '    </a:p>'
                + '  </a:txBody>'
                + '  <a:tcPr'+ cellMargin + cellValign +'>';

            // 5: Borders: Add any borders
            if ( cellOpts.border && typeof cellOpts.border === 'string' ) {
                strXml += '  <a:lnL w="'+ ONEPT +'" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="'+ cellOpts.border +'"/></a:solidFill></a:lnL>';
                strXml += '  <a:lnR w="'+ ONEPT +'" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="'+ cellOpts.border +'"/></a:solidFill></a:lnR>';
                strXml += '  <a:lnT w="'+ ONEPT +'" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="'+ cellOpts.border +'"/></a:solidFill></a:lnT>';
                strXml += '  <a:lnB w="'+ ONEPT +'" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="'+ cellOpts.border +'"/></a:solidFill></a:lnB>';
            }
            else if ( cellOpts.border && Array.isArray(cellOpts.border) ) {
                $.each([ {idx:3,name:'lnL'}, {idx:1,name:'lnR'}, {idx:0,name:'lnT'}, {idx:2,name:'lnB'} ], function(i,obj){
                    if ( cellOpts.border[obj.idx] ) {
                        var strC = '<a:solidFill><a:srgbClr val="'+ ((cellOpts.border[obj.idx].color) ? cellOpts.border[obj.idx].color : '666666') +'"/></a:solidFill>';
                        var intW = (cellOpts.border[obj.idx] && (cellOpts.border[obj.idx].pt || cellOpts.border[obj.idx].pt == 0)) ? (ONEPT * Number(cellOpts.border[obj.idx].pt)) : ONEPT;
                        strXml += '<a:'+ obj.name +' w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strC +'</a:'+ obj.name +'>';
                    }
                    else strXml += '<a:'+ obj.name +' w="0"><a:miter lim="400000" /></a:'+ obj.name +'>';
                });
            }
            else if ( cellOpts.border && typeof cellOpts.border === 'object' ) {
                var intW = (cellOpts.border && (cellOpts.border.pt || cellOpts.border.pt == 0) ) ? (ONEPT * Number(cellOpts.border.pt)) : ONEPT;
                var strClr = '<a:solidFill><a:srgbClr val="'+ ((cellOpts.border.color) ? cellOpts.border.color : '666666') +'"/></a:solidFill>';
                var strAttr = '<a:prstDash val="';
                strAttr += ((cellOpts.border.type && cellOpts.border.type.toLowerCase().indexOf('dash') > -1) ? "sysDash" : "solid" );
                strAttr += '"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>';
                // *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
                strXml += '<a:lnL w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strClr + strAttr +'</a:lnL>';
                strXml += '<a:lnR w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strClr + strAttr +'</a:lnR>';
                strXml += '<a:lnT w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strClr + strAttr +'</a:lnT>';
                strXml += '<a:lnB w="'+ intW +'" cap="flat" cmpd="sng" algn="ctr">'+ strClr + strAttr +'</a:lnB>';
                // *** IMPORTANT! *** LRTB order matters!
            }

            // 6: Close cell Properties & Cell
            strXml += cellFill
                + '  </a:tcPr>'
                + ' </a:tc>';

            // LAST: COLSPAN: Add a 'merged' col for each column being merged (SEE: http://officeopenxml.com/drwTableGrid.php)
            if ( cellOpts.colspan ) {
                for (var tmp=1; tmp<Number(cellOpts.colspan); tmp++) { strXml += '<a:tc hMerge="1"><a:tcPr/></a:tc>'; }
            }
        });

        // B-2: Handle Rowspan as last col case
        // We add dummy cells inside cell loop, but cases where last col is rowspaned
        // by prev row wont be created b/c cell loop above exhausted before the col
        // index of the final col was reached... ANYHOO, add it here when necc.
        if ( arrRowspanCells.filter(function(obj){ return obj.row == rIdx && (obj.col+1) >= $(row).length }).length > 0 ) {
            strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>';
        }

        // C: Complete row
        strXml += '</a:tr>';
    });

    // STEP 5: Complete table
    strXml += '      </a:tbl>'
        + '    </a:graphicData>'
        + '  </a:graphic>'
        + '</p:graphicFrame>';

    // STEP 6: Set table XML
    let strSlideXml = strXml;

    // LAST: Increment counter
    intTableNum++;

    return strSlideXml;

}

