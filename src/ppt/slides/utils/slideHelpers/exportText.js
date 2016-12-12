import { getShapeInfo, genXmlColorSelection, decodeXmlEntities } from '../../utils/helpers';

export  default function ExportText(inSlide, shapeType, slideObj, locationAttr, idx, x, y, cx, cy){

    const ONEPT = 12700, EMU = 914400;
    let strSlideXml = '',
        moreStyles = '',
        moreStylesAttr = '',
        outStyles = '',
        styleData = '';

    // Lines can have zero cy, but text should not
    if ( !slideObj.options.line && cy == 0 ) cy = (EMU * 0.3);

    // Margin/Padding/Inset for textboxes
    if ( slideObj.options.margin && Array.isArray(slideObj.options.margin) ) {
        slideObj.options.bodyProp.lIns = (slideObj.options.margin[0] * ONEPT || 0);
        slideObj.options.bodyProp.rIns = (slideObj.options.margin[1] * ONEPT || 0);
        slideObj.options.bodyProp.bIns = (slideObj.options.margin[2] * ONEPT || 0);
        slideObj.options.bodyProp.tIns = (slideObj.options.margin[3] * ONEPT || 0);
    }
    else if ( (slideObj.options.margin || slideObj.options.margin == 0) && Number.isInteger(slideObj.options.margin) ) {
        slideObj.options.bodyProp.lIns = (slideObj.options.margin * ONEPT);
        slideObj.options.bodyProp.rIns = (slideObj.options.margin * ONEPT);
        slideObj.options.bodyProp.bIns = (slideObj.options.margin * ONEPT);
        slideObj.options.bodyProp.tIns = (slideObj.options.margin * ONEPT);
    }

    var effectsList = '';
    if ( shapeType == null ) shapeType = getShapeInfo(null);

    // A: Start Shape
    strSlideXml = '<p:sp>';

    // B: The addition of the "txBox" attribute is the sole determiner of if an object is a Shape or Textbox
    strSlideXml += `<p:nvSpPr>
                        <p:cNvPr id="${(idx+2)}" name="Object ${(idx+1)}"/>
                        <p:cNvSpPr${(slideObj.options && slideObj.options.isTextBox) ? ' txBox="1"/><p:nvPr/>' : '/><p:nvPr/>'}
                        </p:nvSpPr>
                        <p:spPr>
                            <a:xfrm${locationAttr}>
                                <a:off x="${x}" y="${y}"/>
                                <a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
                            <a:prstGeom prst="${shapeType.name}">
                                <a:avLst/>
                            </a:prstGeom>`;

    if ( slideObj.options ) {
        ( slideObj.options.fill ) ? strSlideXml += genXmlColorSelection(slideObj.options.fill) : strSlideXml += '<a:noFill/>';

        if ( slideObj.options.line ) {
            var lineAttr = '';
            if ( slideObj.options.line_size ) lineAttr += ' w="' + (slideObj.options.line_size * ONEPT) + '"';
            strSlideXml += `<a:ln${lineAttr}>`;
            strSlideXml += genXmlColorSelection( slideObj.options.line );
            if ( slideObj.options.line_head ) strSlideXml += '<a:headEnd type="' + slideObj.options.line_head + '"/>';
            if ( slideObj.options.line_tail ) strSlideXml += '<a:tailEnd type="' + slideObj.options.line_tail + '"/>';
            strSlideXml += '</a:ln>';
        }
    }
    else {
        strSlideXml += '<a:noFill/>';
    }

    // TODO: Implement/document inner/outer-Shadow
    if ( slideObj.options.effects ) {
        for ( var ii = 0, total_size_ii = slideObj.options.effects.length; ii < total_size_ii; ii++ ) {
            switch ( slideObj.options.effects[ii].type ) {
                case 'outerShadow':
                    effectsList += cbGenerateEffects( slideObj.options.effects[ii], 'outerShdw' );
                    break;
                case 'innerShadow':
                    effectsList += cbGenerateEffects( slideObj.options.effects[ii], 'innerShdw' );
                    break;
            }
        }
    }

    if ( effectsList ) strSlideXml += '<a:effectLst>' + effectsList + '</a:effectLst>';

    // TODO 1.5: Text wrapping (copied from MS-PPTX export)
    /*
     // Commented out b/c i'm not even sure this works - current code produces text that wraps in shapes and textboxes, so...
     if ( slideObj.options.textWrap ) {
     strSlideXml += '<a:extLst>'
     + '<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">'
     + '<ma14:wrappingTextBoxFlag xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main" val="1" />'
     + '</a:ext>'
     + '</a:extLst>';
     }
     */

    // B: Close Shape
    strSlideXml += '</p:spPr>';

    if ( slideObj.options ) {
        if ( slideObj.options.align ) {
            switch ( slideObj.options.align ) {
                case 'right':
                    moreStylesAttr += ' algn="r"';
                    break;
                case 'center':
                    moreStylesAttr += ' algn="ctr"';
                    break;
                case 'justify':
                    moreStylesAttr += ' algn="just"';
                    break;
            }
        }

        if ( slideObj.options.indentLevel > 0 ) moreStylesAttr += ' lvl="' + slideObj.options.indentLevel + '"';
    }

    if ( slideObj.options.bullet ) {
        moreStylesAttr = ' marL="228600" indent="-228600"';
        moreStyles = '<a:buSzPct val="100000"/><a:buChar char="&#x2022;"/>';
    }

    if ( moreStyles ) outStyles = '<a:pPr' + moreStylesAttr + '>' + moreStyles + '</a:pPr>';
    else if ( moreStylesAttr ) outStyles = '<a:pPr' + moreStylesAttr + '/>';

    if ( styleData ) strSlideXml += '<p:style>' + styleData + '</p:style>';

    // Bullets: Multi-line bullets must have a complete <a:p><a:pPr></a:pPr><a:r></a:r></a:p> *per line*
    if ( slideObj.options.bullet && typeof slideObj.text == 'string' && slideObj.text.split('\n').length > 0
        && slideObj.text.split('\n')[1] && slideObj.text.split('\n')[1].length > 0 ) {
        strSlideXml += '<p:txBody>' + genXmlBodyProperties( slideObj.options ) + '<a:lstStyle/>';
        $.each(slideObj.text.split('\n'), function(i,line){
            if ( i > 0 ) strSlideXml += '</a:p>';
            strSlideXml += '<a:p>' + outStyles;
            strSlideXml += genXmlTextCommand( slideObj.options, line, inSlide.slide, inSlide.slide.getPageNumber() );
        });
    }
    else if ( typeof slideObj.text == 'string' || typeof slideObj.text == 'number' ) {
        strSlideXml += '<p:txBody>' + genXmlBodyProperties( slideObj.options ) + '<a:lstStyle/><a:p>' + outStyles;
        strSlideXml += genXmlTextCommand( slideObj.options, slideObj.text+'', inSlide.slide, inSlide.slide.getPageNumber() );
    }
    else if ( slideObj.text && slideObj.text.length ) {
        var outBodyOpt = genXmlBodyProperties( slideObj.options );
        strSlideXml += '<p:txBody>' + outBodyOpt + '<a:lstStyle/><a:p>' + outStyles;

        for ( var j = 0, total_size_j = slideObj.text.length; j < total_size_j; j++ ) {
            if ( (typeof slideObj.text[j] == 'object') && slideObj.text[j].text ) {
                strSlideXml += genXmlTextCommand( slideObj.text[j].options || slideObj.options, slideObj.text[j].text, inSlide.slide, outBodyOpt, outStyles, inSlide.slide.getPageNumber() );
            }
            else if ( typeof slideObj.text[j] == 'string' ) {
                strSlideXml += genXmlTextCommand( slideObj.options, slideObj.text[j], inSlide.slide, outBodyOpt, outStyles, inSlide.slide.getPageNumber() );
            }
            else if ( typeof slideObj.text[j] == 'number' ) {
                strSlideXml += genXmlTextCommand( slideObj.options, slideObj.text[j] + '', inSlide.slide, outBodyOpt, outStyles, inSlide.slide.getPageNumber() );
            }
            else if ( (typeof slideObj.text[j] == 'object') && slideObj.text[j].field ) {
                strSlideXml += genXmlTextCommand( slideObj.options, slideObj.text[j], inSlide.slide, outBodyOpt, outStyles, inSlide.slide.getPageNumber() );
            }
        }
    }
    else if ( typeof slideObj.text == 'object' && slideObj.text.field ) {
        strSlideXml += '<p:txBody>' + genXmlBodyProperties( slideObj.options ) + '<a:lstStyle/><a:p>' + outStyles;
        strSlideXml += genXmlTextCommand( slideObj.options, slideObj.text, inSlide.slide, inSlide.slide.getPageNumber() );
    }

    // LAST: End of every paragraph
    if ( typeof slideObj.text !== 'undefined' ) {
        var font_size = '';
        if ( slideObj.options && slideObj.options.font_size ) font_size = ' sz="' + slideObj.options.font_size + '00"';
        strSlideXml += '<a:endParaRPr lang="en-US" '+ font_size +' dirty="0"/></a:p></p:txBody>';
    }

    strSlideXml += (slideObj.type == 'cxn') ? '</p:cxnSp>' : '</p:sp>';

    return strSlideXml;
}

function genXmlBodyProperties( objOptions ) {
    var bodyProperties = '<a:bodyPr';

    if ( objOptions && objOptions.bodyProp ) {
        // A: Enable or disable textwrapping none or square:
        ( objOptions.bodyProp.wrap ) ? bodyProperties += ' wrap="' + objOptions.bodyProp.wrap + '" rtlCol="0"' : bodyProperties += ' wrap="square" rtlCol="0"';

        // B: Set anchorPoints bottom, center or top:
        if ( objOptions.bodyProp.anchor    ) bodyProperties += ' anchor="' + objOptions.bodyProp.anchor + '"';
        if ( objOptions.bodyProp.anchorCtr ) bodyProperties += ' anchorCtr="' + objOptions.bodyProp.anchorCtr + '"';

        // C: Textbox margins [padding]:
        if ( objOptions.bodyProp.bIns || objOptions.bodyProp.bIns == 0 ) bodyProperties += ' bIns="' + objOptions.bodyProp.bIns + '"';
        if ( objOptions.bodyProp.lIns || objOptions.bodyProp.lIns == 0 ) bodyProperties += ' lIns="' + objOptions.bodyProp.lIns + '"';
        if ( objOptions.bodyProp.rIns || objOptions.bodyProp.rIns == 0 ) bodyProperties += ' rIns="' + objOptions.bodyProp.rIns + '"';
        if ( objOptions.bodyProp.tIns || objOptions.bodyProp.tIns == 0 ) bodyProperties += ' tIns="' + objOptions.bodyProp.tIns + '"';

        // D: Close <a:bodyPr element
        bodyProperties += '>';

        // E: NEW: Add auto-fit type tags
        if ( objOptions.shrinkText ) bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000" />'; // MS-PPT > Format Shape > Text Options: "Shrink text on overflow"
        else if ( objOptions.bodyProp.autoFit !== false ) bodyProperties += '<a:spAutoFit/>'; // MS-PPT > Format Shape > Text Options: "Resize shape to fit text"

        // LAST: Close bodyProp
        bodyProperties += '</a:bodyPr>';
    }
    else {
        // DEFAULT:
        bodyProperties += ' wrap="square" rtlCol="0"></a:bodyPr>';
    }

    return bodyProperties;
}


function genXmlTextCommand( text_info, text_string, slide_obj, slide_num ) {
    var area_opt_data = genXmlTextData( text_info, slide_obj );
    var parsedText;
    //var startInfo = '<a:rPr lang="en-US"' + area_opt_data.font_size + area_opt_data.bold + area_opt_data.italic + area_opt_data.underline + area_opt_data.char_spacing + ' dirty="0" smtClean="0"' + (area_opt_data.rpr_info != '' ? ('>' + area_opt_data.rpr_info) : '/>') + '<a:t>';
    var startInfo = '<a:rPr lang="en-US"' + area_opt_data.font_size + area_opt_data.bold  + area_opt_data.underline + area_opt_data.char_spacing + ' dirty="0" smtClean="0"' + (area_opt_data.rpr_info != '' ? ('>' + area_opt_data.rpr_info) : '/>') + '<a:t>';
    var endTag = '</a:r>';
    var outData = '<a:r>' + startInfo;

    if ( text_string.field ) {
        endTag = '</a:fld>';
        var outTextField = pptxFields[text_string.field];
        if ( outTextField === null ) {
            for ( var fieldIntName in pptxFields ) {
                if ( pptxFields[fieldIntName] === text_string.field ) {
                    outTextField = text_string.field;
                    break;
                }
            }

            if ( outTextField === null ) outTextField = 'datetime';
        }

        outData = '<a:fld id="{' + gen_private.plugs.type.msoffice.makeUniqueID ( '5C7A2A3D' ) + '}" type="' + outTextField + '">' + startInfo;
        outData += CreateFieldText( outTextField, slide_num );

    }
    else {
        // Automatic support for newline - split it into multi-p:
        parsedText = text_string.split( "\n" );
        if ( parsedText.length > 1 ) {
            var outTextData = '';
            for ( var i = 0, total_size_i = parsedText.length; i < total_size_i; i++ ) {
                outTextData += outData + decodeXmlEntities(parsedText[i]);

                if ( (i + 1) < total_size_i ) {
                    outTextData += '</a:t></a:r></a:p><a:p>';
                }
            }

            outData = outTextData;

        }
        else {
            outData += text_string.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
        }
    }

    var outBreakP = '';
    if ( text_info.breakLine ) outBreakP += '</a:p><a:p>';

    return outData + '</a:t>' + endTag + outBreakP;
}

function genXmlTextData(text_info, slide_obj) {
    var out_obj = {};

    out_obj.font_size = '';
    out_obj.bold = '';
    out_obj.underline = '';
    out_obj.rpr_info = '';
    out_obj.char_spacing = '';

    if (typeof text_info == 'object') {
        if (text_info.bold) {
            out_obj.bold = ' b="1"';
        }

        if (text_info.underline) {
            out_obj.underline = ' u="sng"';
        }

        if (text_info.font_size) {
            out_obj.font_size = ' sz="' + text_info.font_size + '00"';
        }

        if (text_info.char_spacing) {
            out_obj.char_spacing = ' spc="' + (text_info.char_spacing * 100) + '"';
            // must also disable kerning; otherwise text won't actually expand
            out_obj.char_spacing += ' kern="0"';
        }

        if (text_info.color) {
            out_obj.rpr_info += genXmlColorSelection(text_info.color);
        } else if (slide_obj && slide_obj.color) {
            out_obj.rpr_info += genXmlColorSelection(slide_obj.color);
        }

        if (text_info.font_face) {
            out_obj.rpr_info += '<a:latin typeface="' + text_info.font_face + '" pitchFamily="34" charset="0"/><a:cs typeface="' + text_info.font_face + '" pitchFamily="34" charset="0"/>';
        }
    } else {
        if (slide_obj && slide_obj.color) out_obj.rpr_info += genXmlColorSelection(slide_obj.color);
    }

    if (out_obj.rpr_info != '')
        out_obj.rpr_info += '</a:rPr>';

    return out_obj;
}