import Position from './position'

import {
    inch2Emu,
    getSmartParseNumber,
    encodeXmlEntities,
    genXmlColorSelection
} from '../gen-utils'
import { getSlidesForTableRows } from '../gen-tables'
import {
    ILayout,
    IShadowOptions,
    ISlide,
    ISlideLayout,
    ISlideObject,
    ISlideRel,
    ISlideRelChart,
    ISlideRelMedia,
    ITableCell,
    ITableOptions,
    TableRow,
    ITableCellOpts,
    IObjectOptions,
    IText,
    ITextOpts
} from '../core-interfaces'
import {
    BULLET_TYPES,
    SLIDE_OBJECT_TYPES,
    DEF_CELL_BORDER,
    DEF_CELL_MARGIN_PT,
    DEF_SLIDE_MARGIN_IN,
    ONEPT,
    EMU,
    DEF_FONT_SIZE,
    DEF_FONT_COLOR,
    CRLF
} from '../core-enums'

/**
 * Generate XML Paragraph Properties
 * @param {ISlideObject|IText} textObj - text object
 * @param {boolean} isDefault - array of default relations
 * @return {string} XML
 */
function genXmlParagraphProperties(
    textObj: ISlideObject | IText,
    isDefault: boolean
): string {
    let strXmlBullet = '',
        strXmlLnSpc = '',
        strXmlParaSpc = ''
    let bulletLvl0Margin = 342900
    let tag = isDefault ? 'a:lvl1pPr' : 'a:pPr'

    let paragraphPropXml =
        '<' + tag + (textObj.options.rtlMode ? ' rtl="1" ' : '')

    // A: Build paragraphProperties
    {
        // OPTION: align
        if (textObj.options.align) {
            switch (textObj.options.align) {
                case 'left':
                    paragraphPropXml += ' algn="l"'
                    break
                case 'right':
                    paragraphPropXml += ' algn="r"'
                    break
                case 'center':
                    paragraphPropXml += ' algn="ctr"'
                    break
                case 'justify':
                    paragraphPropXml += ' algn="just"'
                    break
                default:
                    break
            }
        }

        if (textObj.options.lineSpacing) {
            strXmlLnSpc =
                '<a:lnSpc><a:spcPts val="' +
                textObj.options.lineSpacing +
                '00"/></a:lnSpc>'
        }

        // OPTION: indent
        if (
            textObj.options.indentLevel &&
            !isNaN(Number(textObj.options.indentLevel)) &&
            textObj.options.indentLevel > 0
        ) {
            paragraphPropXml += ' lvl="' + textObj.options.indentLevel + '"'
        }

        // OPTION: Paragraph Spacing: Before/After
        if (
            textObj.options.paraSpaceBefore &&
            !isNaN(Number(textObj.options.paraSpaceBefore)) &&
            textObj.options.paraSpaceBefore > 0
        ) {
            strXmlParaSpc +=
                '<a:spcBef><a:spcPts val="' +
                textObj.options.paraSpaceBefore * 100 +
                '"/></a:spcBef>'
        }
        if (
            textObj.options.paraSpaceAfter &&
            !isNaN(Number(textObj.options.paraSpaceAfter)) &&
            textObj.options.paraSpaceAfter > 0
        ) {
            strXmlParaSpc +=
                '<a:spcAft><a:spcPts val="' +
                textObj.options.paraSpaceAfter * 100 +
                '"/></a:spcAft>'
        }

        // OPTION: bullet
        // NOTE: OOXML uses the unicode character set for Bullets
        // EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
        if (typeof textObj.options.bullet === 'object') {
            if (textObj.options.bullet.type) {
                if (
                    textObj.options.bullet.type.toString().toLowerCase() ===
                    'number'
                ) {
                    paragraphPropXml +=
                        ' marL="' +
                        (textObj.options.indentLevel &&
                        textObj.options.indentLevel > 0
                            ? bulletLvl0Margin +
                              bulletLvl0Margin * textObj.options.indentLevel
                            : bulletLvl0Margin) +
                        '" indent="-' +
                        bulletLvl0Margin +
                        '"'
                    strXmlBullet = `<a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="${textObj
                        .options.bullet.style ||
                        'arabicPeriod'}" startAt="${textObj.options.bullet
                        .startAt || '1'}"/>`
                }
            } else if (textObj.options.bullet.code) {
                let bulletCode = '&#x' + textObj.options.bullet.code + ';'

                // Check value for hex-ness (s/b 4 char hex)
                if (
                    /^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.code) ===
                    false
                ) {
                    console.warn(
                        'Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!'
                    )
                    bulletCode = BULLET_TYPES['DEFAULT']
                }

                paragraphPropXml +=
                    ' marL="' +
                    (textObj.options.indentLevel &&
                    textObj.options.indentLevel > 0
                        ? bulletLvl0Margin +
                          bulletLvl0Margin * textObj.options.indentLevel
                        : bulletLvl0Margin) +
                    '" indent="-' +
                    bulletLvl0Margin +
                    '"'
                strXmlBullet =
                    '<a:buSzPct val="100000"/><a:buChar char="' +
                    bulletCode +
                    '"/>'
            }
        } else if (textObj.options.bullet === true) {
            paragraphPropXml +=
                ' marL="' +
                (textObj.options.indentLevel && textObj.options.indentLevel > 0
                    ? bulletLvl0Margin +
                      bulletLvl0Margin * textObj.options.indentLevel
                    : bulletLvl0Margin) +
                '" indent="-' +
                bulletLvl0Margin +
                '"'
            strXmlBullet =
                '<a:buSzPct val="100000"/><a:buChar char="' +
                BULLET_TYPES['DEFAULT'] +
                '"/>'
        } else {
            strXmlBullet = '<a:buNone/>'
        }

        // B: Close Paragraph-Properties
        // IMPORTANT: strXmlLnSpc, strXmlParaSpc, and strXmlBullet require strict ordering - anything out of order is ignored. (PPT-Online, PPT for Mac)
        paragraphPropXml += '>' + strXmlLnSpc + strXmlParaSpc + strXmlBullet
        if (isDefault) {
            paragraphPropXml += genXmlTextRunProperties(textObj.options, true)
        }
        paragraphPropXml += '</' + tag + '>'
    }

    return paragraphPropXml
}

/**
 * Generate XML Text Run Properties (`a:rPr`)
 * @param {IObjectOptions|ITextOpts} opts - text options
 * @param {boolean} isDefault - whether these are the default text run properties
 * @return {string} XML
 */
function genXmlTextRunProperties(
    opts: IObjectOptions | ITextOpts,
    isDefault: boolean
): string {
    let runProps = ''
    let runPropsTag = isDefault ? 'a:defRPr' : 'a:rPr'

    // BEGIN runProperties (ex: `<a:rPr lang="en-US" sz="1600" b="1" dirty="0">`)
    runProps +=
        '<' +
        runPropsTag +
        ' lang="' +
        (opts.lang ? opts.lang : 'en-US') +
        '"' +
        (opts.lang ? ' altLang="en-US"' : '')
    runProps += opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : '' // NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
    runProps += opts.bold ? ' b="1"' : ''
    runProps += opts.italic ? ' i="1"' : ''
    runProps += opts.strike ? ' strike="sngStrike"' : ''
    runProps += opts.underline || opts.hyperlink ? ' u="sng"' : ''
    runProps += opts.subscript
        ? ' baseline="-40000"'
        : opts.superscript
        ? ' baseline="30000"'
        : ''
    runProps += opts.charSpacing
        ? ' spc="' + opts.charSpacing * 100 + '" kern="0"'
        : '' // IMPORTANT: Also disable kerning; otherwise text won't actually expand
    runProps += ' dirty="0">'
    // Color / Font / Outline are children of <a:rPr>, so add them now before closing the runProperties tag
    if (opts.color || opts.fontFace || opts.outline) {
        if (opts.outline && typeof opts.outline === 'object') {
            runProps +=
                '<a:ln w="' +
                Math.round((opts.outline.size || 0.75) * ONEPT) +
                '">' +
                genXmlColorSelection(opts.outline.color || 'FFFFFF') +
                '</a:ln>'
        }
        if (opts.color) runProps += genXmlColorSelection(opts.color)
        if (opts.fontFace) {
            // NOTE: 'cs' = Complex Script, 'ea' = East Asian (use "-120" instead of "0" - per Issue #174); ea must come first (Issue #174)
            runProps +=
                '<a:latin typeface="' +
                opts.fontFace +
                '" pitchFamily="34" charset="0"/>' +
                '<a:ea typeface="' +
                opts.fontFace +
                '" pitchFamily="34" charset="-122"/>' +
                '<a:cs typeface="' +
                opts.fontFace +
                '" pitchFamily="34" charset="-120"/>'
        }
    }

    // Hyperlink support
    if (opts.hyperlink) {
        if (typeof opts.hyperlink !== 'object')
            throw "ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` "
        else if (!opts.hyperlink.url && !opts.hyperlink.slide)
            throw "ERROR: 'hyperlink requires either `url` or `slide`'"
        else if (opts.hyperlink.url) {
            // TODO: (20170410): FUTURE-FEATURE: color (link is always blue in Keynote and PPT online, so usual text run above isnt honored for links..?)
            //runProps += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010! (Issue#74)
            runProps +=
                '<a:hlinkClick r:id="rId' +
                opts.hyperlink.rId +
                '" invalidUrl="" action="" tgtFrame="" tooltip="' +
                (opts.hyperlink.tooltip
                    ? encodeXmlEntities(opts.hyperlink.tooltip)
                    : '') +
                '" history="1" highlightClick="0" endSnd="0"/>'
        } else if (opts.hyperlink.slide) {
            runProps +=
                '<a:hlinkClick r:id="rId' +
                opts.hyperlink.rId +
                '" action="ppaction://hlinksldjump" tooltip="' +
                (opts.hyperlink.tooltip
                    ? encodeXmlEntities(opts.hyperlink.tooltip)
                    : '') +
                '"/>'
        }
    }

    // END runProperties
    runProps += '</' + runPropsTag + '>'

    return runProps
}

/**
 * Builds `<a:r></a:r>` text runs for `<a:p>` paragraphs in textBody
 * @param {IText} textObj - Text object
 * @return {string} XML string
 */
function genXmlTextRun(textObj: IText): string {
    let arrLines = []
    let paraProp = ''
    let xmlTextRun = ''

    // 1: ADD runProperties
    let startInfo = genXmlTextRunProperties(textObj.options, false)

    // 2: LINE-BREAKS/MULTI-LINE: Split text into multi-p:
    arrLines = textObj.text.split(CRLF)
    if (arrLines.length > 1) {
        arrLines.forEach((line, idx) => {
            xmlTextRun +=
                '<a:r>' + startInfo + '<a:t>' + encodeXmlEntities(line)
            // Stop/Start <p>aragraph as long as there is more lines ahead (otherwise its closed at the end of this function)
            if (idx + 1 < arrLines.length)
                xmlTextRun +=
                    (textObj.options.breakLine ? CRLF : '') + '</a:t></a:r>'
        })
    } else {
        // Handle cases where addText `text` was an array of objects - if a text object doesnt contain a '\n' it still need alignment!
        // The first pPr-align is done in makeXml - use line countr to ensure we only add subsequently as needed
        xmlTextRun =
            (textObj.options.align && textObj.options.lineIdx > 0
                ? paraProp
                : '') +
            '<a:r>' +
            startInfo +
            '<a:t>' +
            encodeXmlEntities(textObj.text)
    }

    // Return paragraph with text run
    return xmlTextRun + '</a:t></a:r>'
}

/**
 * Builds `<a:bodyPr></a:bodyPr>` tag for "genXmlTextBody()"
 * @param {ISlideObject | ITableCell} slideObject - various options
 * @return {string} XML string
 */
function genXmlBodyProperties(slideObject: ISlideObject | ITableCell): string {
    let bodyProperties = '<a:bodyPr'

    if (
        (slideObject && slideObject.type === SLIDE_OBJECT_TYPES.text) ||
        (slideObject.type === SLIDE_OBJECT_TYPES.placeholder &&
            slideObject.options.bodyProp)
    ) {
        // PPT-2019 EX: <a:bodyPr wrap="square" lIns="1270" tIns="1270" rIns="1270" bIns="1270" rtlCol="0" anchor="ctr"/>

        // A: Enable or disable textwrapping none or square
        bodyProperties += slideObject.options.bodyProp.wrap
            ? ' wrap="' + slideObject.options.bodyProp.wrap + '"'
            : ' wrap="square"'

        // B: Textbox margins [padding]
        if (
            slideObject.options.bodyProp.lIns ||
            slideObject.options.bodyProp.lIns === 0
        )
            bodyProperties +=
                ' lIns="' + slideObject.options.bodyProp.lIns + '"'
        if (
            slideObject.options.bodyProp.tIns ||
            slideObject.options.bodyProp.tIns === 0
        )
            bodyProperties +=
                ' tIns="' + slideObject.options.bodyProp.tIns + '"'
        if (
            slideObject.options.bodyProp.rIns ||
            slideObject.options.bodyProp.rIns === 0
        )
            bodyProperties +=
                ' rIns="' + slideObject.options.bodyProp.rIns + '"'
        if (
            slideObject.options.bodyProp.bIns ||
            slideObject.options.bodyProp.bIns === 0
        )
            bodyProperties +=
                ' bIns="' + slideObject.options.bodyProp.bIns + '"'

        // C: Add rtl after margins
        bodyProperties += ' rtlCol="0"'

        // D: Add anchorPoints
        if (slideObject.options.bodyProp.anchor)
            bodyProperties +=
                ' anchor="' + slideObject.options.bodyProp.anchor + '"' // VALS: [t,ctr,b]
        if (slideObject.options.bodyProp.vert)
            bodyProperties +=
                ' vert="' + slideObject.options.bodyProp.vert + '"' // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]

        // E: Close <a:bodyPr element
        bodyProperties += '>'

        // F: NEW: Add autofit type tags
        if (slideObject.options.shrinkText)
            bodyProperties +=
                '<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>' // MS-PPT > Format shape > Text Options: "Shrink text on overflow"
        // MS-PPT > Format shape > Text Options: "Resize shape to fit text" [spAutoFit]
        // NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
        bodyProperties +=
            slideObject.options.bodyProp.autoFit !== false
                ? '<a:spAutoFit/>'
                : ''

        // LAST: Close bodyProp
        bodyProperties += '</a:bodyPr>'
    } else {
        // DEFAULT:
        bodyProperties += ' wrap="square" rtlCol="0">'
        bodyProperties += '</a:bodyPr>'
    }

    // LAST: Return Close bodyProp
    return slideObject.type === SLIDE_OBJECT_TYPES.tablecell
        ? '<a:bodyPr/>'
        : bodyProperties
}

/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @note PPT text lines [lines followed by line-breaks] are created using <p>-aragraph's
 * @note Bullets are a paragprah-level formatting device
 * @param {ISlideObject|ITableCell} slideObj - slideObj -OR- table `cell` object
 * @returns XML containing the param object's text and formatting
 */
export function genXmlTextBody(slideObj: ISlideObject | ITableCell): string {
    let opts: IObjectOptions = slideObj.options || {}
    // FIRST: Shapes without text, etc. may be sent here during build, but have no text to render so return an empty string
    if (
        opts &&
        slideObj.type !== SLIDE_OBJECT_TYPES.tablecell &&
        (typeof slideObj.text === 'undefined' || slideObj.text === null)
    )
        return ''

    // Vars
    let arrTextObjects: IText[] = []
    let tagStart =
        slideObj.type === SLIDE_OBJECT_TYPES.tablecell
            ? '<a:txBody>'
            : '<p:txBody>'
    let tagClose =
        slideObj.type === SLIDE_OBJECT_TYPES.tablecell
            ? '</a:txBody>'
            : '</p:txBody>'
    let strSlideXml = tagStart

    // STEP 1: Modify slideObj to be consistent array of `{ text:'', options:{} }`
    /* CASES:
		addText( 'string' )
		addText( 'line1\n line2' )
		addText( ['barry','allen'] )
		addText( [{text'word1'}, {text:'word2'}] )
		addText( [{text'line1\n line2'}, {text:'end word'}] )
	*/
    // A: Transform string/number into complex object
    if (
        typeof slideObj.text === 'string' ||
        typeof slideObj.text === 'number'
    ) {
        slideObj.text = [
            { text: slideObj.text.toString(), options: opts || {} }
        ]
    }

    // STEP 2: Grab options, format line-breaks, etc.
    if (Array.isArray(slideObj.text)) {
        slideObj.text.forEach((obj, idx) => {
            // A: Set options
            obj.options = obj.options || opts || {}
            if (idx === 0 && obj.options && !obj.options.bullet && opts.bullet)
                obj.options.bullet = opts.bullet

            // B: Cast to text-object and fix line-breaks (if needed)
            if (typeof obj.text === 'string' || typeof obj.text === 'number') {
                // 1: Convert "\n" or any variation into CRLF
                obj.text = obj.text.toString().replace(/\r*\n/g, CRLF)

                // 2: Handle strings that contain "\n"
                if (obj.text.indexOf(CRLF) > -1) {
                    // Remove trailing linebreak (if any) so the "if" below doesnt create a double CRLF+CRLF line ending!
                    obj.text = obj.text.replace(/\r\n$/g, '')
                    // Plain strings like "hello \n world" or "first line\n" need to have lineBreaks set to become 2 separate lines as intended
                    obj.options.breakLine = true
                }

                // 3: Add CRLF line ending if `breakLine`
                if (
                    obj.options.breakLine &&
                    !obj.options.bullet &&
                    !obj.options.align &&
                    idx + 1 < slideObj.text.length
                )
                    obj.text += CRLF
            }

            // C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
            if (obj.options.breakLine || obj.text.indexOf(CRLF) > -1) {
                obj.text.split(CRLF).forEach((line, lineIdx) => {
                    // Add line-breaks if not bullets/aligned (we add CRLF for those below in STEP 3)
                    // NOTE: Use "idx>0" so lines wont start with linebreak (eg:empty first line)
                    arrTextObjects.push({
                        text:
                            (lineIdx > 0 &&
                            obj.options.breakLine &&
                            !obj.options.bullet &&
                            !obj.options.align
                                ? CRLF
                                : '') + line,
                        options: obj.options
                    })
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
        strSlideXml += genXmlBodyProperties(slideObj)

        // B: 'lstStyle'
        // NOTE: shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
        // FIXME: LINE horiz-align doesnt work (text is always to the left inside line) (FYI: the PPT code diff is substantial!)
        if (opts.h === 0 && opts.line && opts.align) {
            strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>'
        } else if (slideObj.type === 'placeholder') {
            strSlideXml += `<a:lstStyle>${genXmlParagraphProperties(
                slideObj,
                true
            )}</a:lstStyle>`
        } else {
            strSlideXml += '<a:lstStyle/>'
        }
    }

    // STEP 4: Loop over each text object and create paragraph props, text run, etc.
    arrTextObjects.forEach((textObj, idx) => {
        // Clear/Increment loop vars
        let paragraphPropXml =
            '<a:pPr ' + (textObj.options.rtlMode ? ' rtl="1" ' : '')
        textObj.options.lineIdx = idx

        // A: Inherit pPr-type options from parent shape's `options`
        textObj.options.align = textObj.options.align || opts.align
        textObj.options.lineSpacing =
            textObj.options.lineSpacing || opts.lineSpacing
        textObj.options.indentLevel =
            textObj.options.indentLevel || opts.indentLevel
        textObj.options.paraSpaceBefore =
            textObj.options.paraSpaceBefore || opts.paraSpaceBefore
        textObj.options.paraSpaceAfter =
            textObj.options.paraSpaceAfter || opts.paraSpaceAfter

        textObj.options.lineIdx = idx
        paragraphPropXml = genXmlParagraphProperties(textObj, false)

        // B: Start paragraph if this is the first text obj, or if current textObj is about to be bulleted or aligned
        if (idx === 0) {
            // Add paragraphProperties right after <p> before textrun(s) begin
            strSlideXml += '<a:p>' + paragraphPropXml
        } else if (
            idx > 0 &&
            (typeof textObj.options.bullet !== 'undefined' ||
                typeof textObj.options.align !== 'undefined')
        ) {
            strSlideXml += '</a:p><a:p>' + paragraphPropXml
        }

        // C: Inherit any main options (color, fontSize, etc.)
        // We only pass the text.options to genXmlTextRun (not the Slide.options),
        // so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
        Object.entries(opts).forEach(([key, val]) => {
            // NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
            if (key !== 'bullet' && !textObj.options[key])
                textObj.options[key] = val
        })

        // D: Add formatted textrun
        strSlideXml += genXmlTextRun(textObj)
    })

    // STEP 5: Append 'endParaRPr' (when needed) and close current open paragraph
    // NOTE: (ISSUE#20, ISSUE#193): Add 'endParaRPr' with font/size props or PPT default (Arial/18pt en-us) is used making row "too tall"/not honoring options
    if (
        slideObj.type === SLIDE_OBJECT_TYPES.tablecell &&
        (opts.fontSize || opts.fontFace)
    ) {
        if (opts.fontFace) {
            strSlideXml += strSlideXml += `
			<a:endParaRPr lang="${opts.lang ? opts.lang : 'en-US'}"${
                opts.fontSize ? ` sz="${Math.round(opts.fontSize)}00"` : ''
            } dirty="0">
              <a:latin typeface="${opts.fontFace}" charset="0"/>
			  <a:ea typeface="${opts.fontFace}" charset="0"/>
			  <a:cs typeface="${opts.fontFace}" charset="0"/>
			</a:endParaRPr>`
        } else {
            strSlideXml += `<a:endParaRPr lang="${
                opts.lang ? opts.lang : 'en-US'
            }"${
                opts.fontSize ? ` sz="${Math.round(opts.fontSize)}00"` : ''
            } dirty="0"/>`
        }
    } else {
        strSlideXml +=
            '<a:endParaRPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>' // NOTE: Added 20180101 to address PPT-2007 issues
    }
    strSlideXml += '</a:p>'

    // STEP 6: Close the textBody
    strSlideXml += tagClose

    // LAST: Return XML
    return strSlideXml
}

function createHyperlinkRels(
    target: ISlide,
    text: number | string | IText | IText[] | ITableCell[][]
) {
    let textObjs = []

    // Only text objects can have hyperlinks, bail when text param is plain text
    if (typeof text === 'string' || typeof text === 'number') return
    // IMPORTANT: "else if" Array.isArray must come before typeof===object! Otherwise, code will exhaust recursion!
    else if (Array.isArray(text)) textObjs = text
    else if (typeof text === 'object') textObjs = [text]

    textObjs.forEach((text: IText) => {
        // `text` can be an array of other `text` objects (table cell word-level formatting), continue parsing using recursion
        if (Array.isArray(text)) createHyperlinkRels(target, text)
        else if (
            text &&
            typeof text === 'object' &&
            text.options &&
            text.options.hyperlink &&
            !text.options.hyperlink.rId
        ) {
            if (typeof text.options.hyperlink !== 'object')
                console.log(
                    "ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:'https://github.com'}` "
                )
            else if (
                !text.options.hyperlink.url &&
                !text.options.hyperlink.slide
            )
                console.log(
                    "ERROR: 'hyperlink requires either: `url` or `slide`'"
                )
            else {
                let relId =
                    target.rels.length +
                    target.relsChart.length +
                    target.relsMedia.length +
                    1

                target.rels.push({
                    type: SLIDE_OBJECT_TYPES.hyperlink,
                    data: text.options.hyperlink.slide ? 'slide' : 'dummy',
                    rId: relId,
                    Target:
                        encodeXmlEntities(text.options.hyperlink.url) ||
                        text.options.hyperlink.slide.toString()
                })

                text.options.hyperlink.rId = relId
            }
        }
    })
}

export default class TableElement {
    type = SLIDE_OBJECT_TYPES.newtext

    arrTabRows
    options
    position: Position

    constructor(
        target: ISlide,
        tableRows: TableRow[],
        options: ITableOptions,
        slideLayout: ISlideLayout,
        presLayout: ILayout,
        addSlide: Function,
        getSlide: Function
    ) {
        let opt: ITableOptions =
            options && typeof options === 'object' ? options : {}
        let slides: ISlide[] = [target] // Create array of Slides as more may be added by auto-paging

        // STEP 1: REALITY-CHECK
        {
            // A: check for empty
            if (
                tableRows === null ||
                tableRows.length === 0 ||
                !Array.isArray(tableRows)
            ) {
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
                        options: typeof cell === 'object' ? cell.options : null
                    }
                    if (typeof cell === 'string' || typeof cell === 'number')
                        newCell.text = cell.toString()
                    else if (cell.text) {
                        // Cell can contain complex text type, or string, or number
                        if (
                            typeof cell.text === 'string' ||
                            typeof cell.text === 'number'
                        )
                            newCell.text = cell.text.toString()
                        else if (cell.text) newCell.text = cell.text
                        // Capture options
                        if (cell.options) newCell.options = cell.options
                    }
                    newRow.push(newCell)
                })
            } else {
                console.log(
                    'addTable: tableRows has a bad row. A row should be an array of cells. You provided:'
                )
                console.log(row)
            }

            arrRows.push(newRow)
        })

        // STEP 3: Set options
        opt.x = getSmartParseNumber(
            opt.x || (opt.x === 0 ? 0 : EMU / 2),
            'X',
            presLayout
        )
        opt.y = getSmartParseNumber(
            opt.y || (opt.y === 0 ? 0 : EMU / 2),
            'Y',
            presLayout
        )
        if (opt.h) opt.h = getSmartParseNumber(opt.h, 'Y', presLayout) // NOTE: Dont set default `h` - leaving it null triggers auto-rowH in `makeXMLSlide()`
        opt.autoPage = typeof opt.autoPage === 'boolean' ? opt.autoPage : false
        opt.fontSize = opt.fontSize || DEF_FONT_SIZE
        opt.autoPageLineWeight =
            typeof opt.autoPageLineWeight !== 'undefined' &&
            !isNaN(Number(opt.autoPageLineWeight))
                ? Number(opt.autoPageLineWeight)
                : 0
        opt.margin =
            opt.margin === 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_PT
        if (typeof opt.margin === 'number')
            opt.margin = [
                Number(opt.margin),
                Number(opt.margin),
                Number(opt.margin),
                Number(opt.margin)
            ]
        if (opt.autoPageLineWeight > 1) opt.autoPageLineWeight = 1
        else if (opt.autoPageLineWeight < -1) opt.autoPageLineWeight = -1
        // Set default color if needed (table option > inherit from Slide > default to black)
        if (!opt.color) opt.color = opt.color || DEF_FONT_COLOR

        // Set/Calc table width
        // Get slide margins - start with default values, then adjust if master or slide margins exist
        let arrTableMargin = DEF_SLIDE_MARGIN_IN
        // Case 1: Master margins
        if (slideLayout && typeof slideLayout.margin !== 'undefined') {
            if (Array.isArray(slideLayout.margin))
                arrTableMargin = slideLayout.margin
            else if (!isNaN(Number(slideLayout.margin)))
                arrTableMargin = [
                    Number(slideLayout.margin),
                    Number(slideLayout.margin),
                    Number(slideLayout.margin),
                    Number(slideLayout.margin)
                ]
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
            } else if (
                opt.colW &&
                Array.isArray(opt.colW) &&
                opt.colW.length !== arrRows[0].length
            ) {
                console.warn(
                    'addTable: colW.length != data.length! Defaulting to evenly distributed col widths.'
                )

                let numColWidth = Math.floor(
                    (presLayout.width / EMU -
                        arrTableMargin[1] -
                        arrTableMargin[3]) /
                        arrRows[0].length
                )
                opt.colW = []
                for (let idx = 0; idx < arrRows[0].length; idx++) {
                    opt.colW.push(numColWidth)
                }
                opt.w = Math.floor(numColWidth * arrRows[0].length)
            }
        } else {
            opt.w = Math.floor(
                presLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3]
            )
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
                    row[idy] = {
                        type: SLIDE_OBJECT_TYPES.tablecell,
                        text: row[idy].toString(),
                        options: opt
                    }
                } else if (typeof cell === 'object') {
                    // ARG0: `text`
                    if (typeof cell.text === 'number')
                        row[idy].text = row[idy].text.toString()
                    else if (
                        typeof cell.text === 'undefined' ||
                        cell.text === null
                    )
                        row[idy].text = ''

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

        this.position = new Position({
            x: opt.x,
            y: opt.y,
            w: opt.w,
            h: opt.h
        })
        // STEP 6: Auto-Paging: (via {options} and used internally)
        // (used internally by `tableToSlides()` to not engage recursion - we've already paged the table data, just add this one)
        if (opt && opt.autoPage === false) {
            // Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
            createHyperlinkRels(target, arrRows)

            // Add data (NOTE: Use `extend` to avoid mutation)
            this.arrTabRows = arrRows
            this.options = opt
        } else {
            console.error('auto paging disabled for now')
            /*
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
			*/
        }
    }

    render(idx, presLayout) {
        let objTableGrid = {}
        let arrTabRows = this.arrTabRows
        let objTabOpts = this.options
        let intColCnt = 0,
            intColW = 0
        let cellOpts: ITableCellOpts

        // Calc number of columns
        // NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
        // ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
        arrTabRows[0].forEach(cell => {
            cellOpts = cell.options || null
            intColCnt +=
                cellOpts && cellOpts.colspan ? Number(cellOpts.colspan) : 1
        })

        // STEP 1: Start Table XML
        // NOTE: Non-numeric cNvPr id values will trigger "presentation needs repair" type warning in MS-PPT-2013
        let strXml = [
            '<p:graphicFrame>',
            '  <p:nvGraphicFramePr>',
            `    <p:cNvPr id="${idx + 1}" name="Table ${idx}"/>`,
            '    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>',
            '    <p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>',
            '  </p:nvGraphicFramePr>',
            this.position.render(presLayout),
            '  <a:graphic>',
            '    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">',
            '      <a:tbl>',
            '        <a:tblPr/>'
        ].join('')

        // + '        <a:tblPr bandRow="1"/>';
        // TODO: Support banded rows, first/last row, etc.
        // NOTE: Banding, etc. only shows when using a table style! (or set alt row color if banding)
        // <a:tblPr firstCol="0" firstRow="0" lastCol="0" lastRow="0" bandCol="0" bandRow="1">

        // STEP 2: Set column widths
        // Evenly distribute cols/rows across size provided when applicable (calc them if only overall dimensions were provided)
        // A: Col widths provided?
        if (Array.isArray(objTabOpts.colW)) {
            strXml += '<a:tblGrid>'
            for (let col = 0; col < intColCnt; col++) {
                strXml +=
                    '<a:gridCol w="' +
                    Math.round(
                        inch2Emu(objTabOpts.colW[col]) ||
                            (typeof this.options.w === 'number'
                                ? this.options.w
                                : 1) / intColCnt
                    ) +
                    '"/>'
            }
            strXml += '</a:tblGrid>'
        }
        // B: Table Width provided without colW? Then distribute cols
        else {
            intColW = objTabOpts.colW ? objTabOpts.colW : EMU
            if (this.options.w && !objTabOpts.colW)
                intColW = Math.round(
                    (typeof this.options.w === 'number' ? this.options.w : 1) /
                        intColCnt
                )
            strXml += '<a:tblGrid>'
            for (let col = 0; col < intColCnt; col++) {
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
        /*
					Object ex: key = rowIdx / val = [cells] cellIdx { 0:{type: "tablecell", text: Array(1), options: {…}}, 1:... }
					{0: {…}, 1: {…}, 2: {…}, 3: {…}}
				*/
        arrTabRows.forEach((row, rIdx) => {
            // A: Create row if needed (recall one may be created in loop below for rowspans, so dont assume we need to create one each iteration)
            if (!objTableGrid[rIdx]) objTableGrid[rIdx] = {}

            // B: Loop over all cells
            row.forEach((cell, cIdx) => {
                // DESIGN: NOTE: Row cell arrays can be "uneven" (diff cell count in each) due to rowspan/colspan
                // Therefore, for each cell we run 0->colCount to determine the correct slot for it to reside
                // as the uneven/mixed nature of the data means we cannot use the cIdx value alone.
                // E.g.: the 2nd element in the row array may actually go into the 5th table grid row cell b/c of colspans!
                for (let idx = 0; cIdx + idx < intColCnt; idx++) {
                    let currColIdx = cIdx + idx

                    if (!objTableGrid[rIdx][currColIdx]) {
                        // A: Set this cell
                        objTableGrid[rIdx][currColIdx] = cell

                        // B: Handle `colspan` or `rowspan` (a {cell} cant have both! TODO: FUTURE: ROWSPAN & COLSPAN in same cell)
                        if (
                            cell &&
                            cell.options &&
                            cell.options.colspan &&
                            !isNaN(Number(cell.options.colspan))
                        ) {
                            for (
                                let idy = 1;
                                idy < Number(cell.options.colspan);
                                idy++
                            ) {
                                objTableGrid[rIdx][currColIdx + idy] = {
                                    hmerge: true,
                                    text: 'hmerge'
                                }
                            }
                        } else if (
                            cell &&
                            cell.options &&
                            cell.options.rowspan &&
                            !isNaN(Number(cell.options.rowspan))
                        ) {
                            for (
                                let idz = 1;
                                idz < Number(cell.options.rowspan);
                                idz++
                            ) {
                                if (!objTableGrid[rIdx + idz])
                                    objTableGrid[rIdx + idz] = {}
                                objTableGrid[rIdx + idz][currColIdx] = {
                                    vmerge: true,
                                    text: 'vmerge'
                                }
                            }
                        }

                        // C: Break out of colCnt loop now that slot has been filled
                        break
                    }
                }
            })
        })

        /* DEBUG: useful for rowspan/colspan testing
				if ( objTabOpts.verbose ) {
					console.table(objTableGrid);
					let arrText = [];
					objTableGrid.forEach(function(row){ let arrRow = []; row.forEach(row,function(cell){ arrRow.push(cell.text); }); arrText.push(arrRow); });
					console.table( arrText );
				}
				*/

        // STEP 4: Build table rows/cells
        Object.entries(objTableGrid).forEach(([rIdx, rowObj]) => {
            // A: Table Height provided without rowH? Then distribute rows
            let intRowH = 0 // IMPORTANT: Default must be zero for auto-sizing to work
            if (Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx])
                intRowH = inch2Emu(Number(objTabOpts.rowH[rIdx]))
            else if (objTabOpts.rowH && !isNaN(Number(objTabOpts.rowH)))
                intRowH = inch2Emu(Number(objTabOpts.rowH))
            else if (this.options.cy || this.options.h)
                intRowH =
                    (this.options.h
                        ? inch2Emu(this.options.h)
                        : typeof this.options.cy === 'number'
                        ? this.options.cy
                        : 1) / arrTabRows.length

            // B: Start row
            strXml += '<a:tr h="' + intRowH + '">'

            // C: Loop over each CELL
            Object.entries(rowObj).forEach(([_cIdx, cellObj]) => {
                let cell: ITableCell = cellObj

                // 1: "hmerge" cells are just place-holders in the table grid - skip those and go to next cell
                if (cell.hmerge) return

                // 2: OPTIONS: Build/set cell options
                let cellOpts = cell.options || ({} as ITableCell['options'])
                cell.options = cellOpts

                // B: Inherit some options from table when cell options dont exist
                // @see: http://officeopenxml.com/drwTableCellProperties-alignment.php
                ;[
                    'align',
                    'bold',
                    'border',
                    'color',
                    'fill',
                    'fontFace',
                    'fontSize',
                    'margin',
                    'underline',
                    'valign'
                ].forEach(name => {
                    if (
                        objTabOpts[name] &&
                        !cellOpts[name] &&
                        cellOpts[name] !== 0
                    )
                        cellOpts[name] = objTabOpts[name]
                })

                let cellValign = cellOpts.valign
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
                let cellColspan = cellOpts.colspan
                    ? ' gridSpan="' + cellOpts.colspan + '"'
                    : ''
                let cellRowspan = cellOpts.rowspan
                    ? ' rowSpan="' + cellOpts.rowspan + '"'
                    : ''
                let cellFill =
                    (cell.optImp && cell.optImp.fill) || cellOpts.fill
                        ? ' <a:solidFill><a:srgbClr val="' +
                          (
                              (cell.optImp && cell.optImp.fill) ||
                              (typeof cellOpts.fill === 'string'
                                  ? cellOpts.fill.replace('#', '')
                                  : '')
                          ).toUpperCase() +
                          '"/></a:solidFill>'
                        : ''
                let cellMargin =
                    cellOpts.margin === 0 || cellOpts.margin
                        ? cellOpts.margin
                        : DEF_CELL_MARGIN_PT
                if (
                    !Array.isArray(cellMargin) &&
                    typeof cellMargin === 'number'
                )
                    cellMargin = [
                        cellMargin,
                        cellMargin,
                        cellMargin,
                        cellMargin
                    ]
                let cellMarginXml =
                    ' marL="' +
                    cellMargin[3] * ONEPT +
                    '" marR="' +
                    cellMargin[1] * ONEPT +
                    '" marT="' +
                    cellMargin[0] * ONEPT +
                    '" marB="' +
                    cellMargin[2] * ONEPT +
                    '"'

                // TODO: Cell NOWRAP property (text wrap: add to a:tcPr (horzOverflow="overflow" or whatever options exist)

                // 3: ROWSPAN: Add dummy cells for any active rowspan
                if (cell.vmerge) {
                    strXml += '<a:tc vMerge="1"><a:tcPr/></a:tc>'
                    return
                }

                // 4: Set CELL content and properties ==================================
                strXml +=
                    '<a:tc' +
                    cellColspan +
                    cellRowspan +
                    '>' +
                    genXmlTextBody(cell) +
                    '<a:tcPr' +
                    cellMarginXml +
                    cellValign +
                    '>'

                // 5: Borders: Add any borders
                if (
                    cellOpts.border &&
                    !Array.isArray(cellOpts.border) &&
                    cellOpts.border.type === 'none'
                ) {
                    strXml +=
                        '  <a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL>'
                    strXml +=
                        '  <a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR>'
                    strXml +=
                        '  <a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT>'
                    strXml +=
                        '  <a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>'
                } else if (
                    cellOpts.border &&
                    typeof cellOpts.border === 'string'
                ) {
                    strXml +=
                        '  <a:lnL w="' +
                        ONEPT +
                        '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' +
                        cellOpts.border +
                        '"/></a:solidFill></a:lnL>'
                    strXml +=
                        '  <a:lnR w="' +
                        ONEPT +
                        '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' +
                        cellOpts.border +
                        '"/></a:solidFill></a:lnR>'
                    strXml +=
                        '  <a:lnT w="' +
                        ONEPT +
                        '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' +
                        cellOpts.border +
                        '"/></a:solidFill></a:lnT>'
                    strXml +=
                        '  <a:lnB w="' +
                        ONEPT +
                        '" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="' +
                        cellOpts.border +
                        '"/></a:solidFill></a:lnB>'
                } else if (cellOpts.border && Array.isArray(cellOpts.border)) {
                    ;[
                        { idx: 3, name: 'lnL' },
                        { idx: 1, name: 'lnR' },
                        { idx: 0, name: 'lnT' },
                        { idx: 2, name: 'lnB' }
                    ].forEach(obj => {
                        if (cellOpts.border[obj.idx]) {
                            let strC =
                                '<a:solidFill><a:srgbClr val="' +
                                (cellOpts.border[obj.idx].color
                                    ? cellOpts.border[obj.idx].color
                                    : DEF_CELL_BORDER.color) +
                                '"/></a:solidFill>'
                            let intW =
                                cellOpts.border[obj.idx] &&
                                (cellOpts.border[obj.idx].pt ||
                                    cellOpts.border[obj.idx].pt === 0)
                                    ? ONEPT *
                                      Number(cellOpts.border[obj.idx].pt)
                                    : ONEPT
                            strXml +=
                                '<a:' +
                                obj.name +
                                ' w="' +
                                intW +
                                '" cap="flat" cmpd="sng" algn="ctr">' +
                                strC +
                                '</a:' +
                                obj.name +
                                '>'
                        } else
                            strXml +=
                                '<a:' +
                                obj.name +
                                ' w="0"><a:miter lim="400000"/></a:' +
                                obj.name +
                                '>'
                    })
                } else if (cellOpts.border && !Array.isArray(cellOpts.border)) {
                    let intW =
                        cellOpts.border &&
                        (cellOpts.border.pt || cellOpts.border.pt === 0)
                            ? ONEPT * Number(cellOpts.border.pt)
                            : ONEPT
                    let strClr =
                        '<a:solidFill><a:srgbClr val="' +
                        (cellOpts.border.color
                            ? cellOpts.border.color.replace('#', '')
                            : DEF_CELL_BORDER.color) +
                        '"/></a:solidFill>'
                    let strAttr = '<a:prstDash val="'
                    strAttr +=
                        cellOpts.border.type &&
                        cellOpts.border.type.toLowerCase().indexOf('dash') > -1
                            ? 'sysDash'
                            : 'solid'
                    strAttr +=
                        '"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>'
                    // *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
                    strXml +=
                        '<a:lnL w="' +
                        intW +
                        '" cap="flat" cmpd="sng" algn="ctr">' +
                        strClr +
                        strAttr +
                        '</a:lnL>'
                    strXml +=
                        '<a:lnR w="' +
                        intW +
                        '" cap="flat" cmpd="sng" algn="ctr">' +
                        strClr +
                        strAttr +
                        '</a:lnR>'
                    strXml +=
                        '<a:lnT w="' +
                        intW +
                        '" cap="flat" cmpd="sng" algn="ctr">' +
                        strClr +
                        strAttr +
                        '</a:lnT>'
                    strXml +=
                        '<a:lnB w="' +
                        intW +
                        '" cap="flat" cmpd="sng" algn="ctr">' +
                        strClr +
                        strAttr +
                        '</a:lnB>'
                    // *** IMPORTANT! *** LRTB order matters!
                }

                // 6: Close cell Properties & Cell
                strXml += cellFill
                strXml += '  </a:tcPr>'
                strXml += ' </a:tc>'

                // LAST: COLSPAN: Add a 'merged' col for each column being merged (SEE: http://officeopenxml.com/drwTableGrid.php)
                if (cellOpts.colspan) {
                    for (let tmp = 1; tmp < Number(cellOpts.colspan); tmp++) {
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
        return strXml
    }
}
