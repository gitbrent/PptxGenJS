/**
 * PptxGenJS: XML Generation
 */

import {
	BULLET_TYPES,
	CRLF,
	DEF_FONT_SIZE,
	DEF_SLIDE_MARGIN_IN,
	EMU,
	LAYOUT_IDX_SERIES_BASE,
	ONEPT,
	PLACEHOLDER_TYPES,
	SLDNUMFLDID,
} from './enums'
import { ISlide, IShadowOpts, ILayout, ISlideLayout, ITableCell } from './interfaces'
import { encodeXmlEntities, inch2Emu, genXmlColorSelection } from './utils'
import { gObjPptxShapes } from './lib-shapes'
import { slideObjectToXml, slideObjectRelationsToXml } from './gen-objects'

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
	if (slideObj.options && !slideObj.options.isTableCell && (typeof slideObj.text === 'undefined' || slideObj.text == null)) return ''

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

/**
 * Magic happens here
 */
function parseTextToLines(cell: ITableCell, inWidth: number): Array<string> {
	var CHAR = 2.2 + (cell.opts && cell.opts.lineWeight ? cell.opts.lineWeight : 0) // Character Constant (An approximation of the Golden Ratio)
	var CPL = (inWidth * EMU) / ((cell.opts.fontSize || DEF_FONT_SIZE) / CHAR) // Chars-Per-Line
	var arrLines = []
	var strCurrLine = ''

	// Allow a single space/whitespace as cell text
	if (cell.text && cell.text.trim() == '') return [' ']

	// A: Remove leading/trailing space
	var inStr = (cell.text || '').toString().trim()

	// B: Build line array
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
		if (strCurrLine) arrLines.push(jQuery.trim(strCurrLine) + CRLF)
		strCurrLine = ''
	})

	// C: Remove trailing linebreak
	arrLines[arrLines.length - 1] = jQuery.trim(arrLines[arrLines.length - 1])

	// D: Return lines
	return arrLines
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

export function genXmlPlaceholder(placeholderObj) {
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
	strXml += slideObjectToXml(objSlide)
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
		objSlide.number +
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
	strXml += slideObjectToXml(objSlideLayout as ISlideLayout)
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
	strXml += slideObjectToXml(objSlide)
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
	return slideObjectRelationsToXml(slideLayouts[layoutNumber - 1], [
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
	return slideObjectRelationsToXml(slides[slideNumber - 1], [
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

	return slideObjectRelationsToXml(masterSlideObject, defaultRels)
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
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {IShadowOpts} `shadowOpts`
 */
export function correctShadowOptions(shadowOpts: IShadowOpts) {
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

export function correctGridLineOptions(glOpts) {
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

export function getShapeInfo(shapeName) {
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

export function getSlidesForTableRows(inArrRows, opts, presLayout: ILayout) {
	var LINEH_MODIFIER = 1.9
	var opts = opts || {}
	var arrInchMargins = DEF_SLIDE_MARGIN_IN // (0.5" on all sides)
	var arrObjTabHeadRows = opts.arrObjTabHeadRows || []
	var arrObjSlides = [],
		arrRows = [],
		currRow = [],
		numCols = 0
	var emuTabCurrH = 0,
		emuSlideTabW = EMU * 1,
		emuSlideTabH = EMU * 1

	if (opts.debug) console.log('------------------------------------')
	if (opts.debug) console.log('opts.w ............. = ' + (opts.w || '').toString())
	if (opts.debug) console.log('opts.colW .......... = ' + (opts.colW || '').toString())
	if (opts.debug) console.log('opts.slideMargin ... = ' + (opts.slideMargin || '').toString())

	// NOTE: Use default size as zero cell margin is causing our tables to be too large and touch bottom of slide!
	if (!opts.slideMargin && opts.slideMargin != 0) opts.slideMargin = DEF_SLIDE_MARGIN_IN[0]

	// STEP 1: Calc margins/usable space
	if (opts.slideMargin || opts.slideMargin == 0) {
		if (Array.isArray(opts.slideMargin)) arrInchMargins = opts.slideMargin
		else if (!isNaN(opts.slideMargin)) arrInchMargins = [opts.slideMargin, opts.slideMargin, opts.slideMargin, opts.slideMargin]
	} else if (opts && opts.master && opts.master.margin) {
		if (Array.isArray(opts.master.margin)) arrInchMargins = opts.master.margin
		else if (!isNaN(opts.master.margin)) arrInchMargins = [opts.master.margin, opts.master.margin, opts.master.margin, opts.master.margin]
	}

	// STEP 2: Calc number of columns
	// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
	// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
	inArrRows[0].forEach(cell => {
		if (!cell) cell = {}
		var cellOpts = cell.options || cell.opts || null
		numCols += cellOpts && cellOpts.colspan ? cellOpts.colspan : 1
	})

	if (opts.debug) console.log('arrInchMargins ..... = ' + arrInchMargins.toString())
	if (opts.debug) console.log('numCols ............ = ' + numCols)

	// Calc opts.w if we can
	if (!opts.w && opts.colW) {
		if (Array.isArray(opts.colW))
			opts.colW.forEach(val => {
				opts.w += val
			})
		else {
			opts.w = opts.colW * numCols
		}
	}

	// STEP 2: Calc usable space/table size now that we have usable space calc'd
	emuSlideTabW = opts.w ? inch2Emu(opts.w) : presLayout.width - inch2Emu((opts.x || arrInchMargins[1]) + arrInchMargins[3])
	if (opts.debug) console.log('emuSlideTabW (in) ........ = ' + (emuSlideTabW / EMU).toFixed(1))
	if (opts.debug) console.log('presLayout.h ..... = ' + presLayout.height / EMU)

	// STEP 3: Calc column widths if needed so we can subsequently calc lines (we need `emuSlideTabW`!)
	if (!opts.colW || !Array.isArray(opts.colW)) {
		if (opts.colW && !isNaN(Number(opts.colW))) {
			var arrColW = []
			inArrRows[0].forEach(() => {
				arrColW.push(opts.colW)
			})
			opts.colW = []
			arrColW.forEach(val => {
				opts.colW.push(val)
			})
		}
		// No column widths provided? Then distribute cols.
		else {
			opts.colW = []
			for (var iCol = 0; iCol < numCols; iCol++) {
				opts.colW.push(emuSlideTabW / EMU / numCols)
			}
		}
	}

	// STEP 4: Iterate over each line and perform magic =========================
	// NOTE: inArrRows will be an array of {text:'', opts{}} whether from `addSlidesForTable()` or `.addTable()`
	inArrRows.forEach((row, iRow) => {
		// A: Reset ROW variables
		var arrCellsLines = [],
			arrCellsLineHeights = [],
			emuRowH = 0,
			intMaxLineCnt = 0,
			intMaxColIdx = 0

		// B: Calc usable vertical space/table height
		// NOTE: Use margins after the first Slide (dont re-use opt.y - it could've been halfway down the page!) (ISSUE#43,ISSUE#47,ISSUE#48)
		if (arrObjSlides.length > 0) {
			emuSlideTabH = presLayout.height - inch2Emu((opts.y / EMU < arrInchMargins[0] ? opts.y / EMU : arrInchMargins[0]) + arrInchMargins[2])
			// Use whichever is greater: area between margins or the table H provided (dont shrink usable area - the whole point of over-riding X on paging is to *increarse* usable space)
			if (emuSlideTabH < opts.h) emuSlideTabH = opts.h
		} else emuSlideTabH = opts.h ? opts.h : presLayout.height - inch2Emu((opts.y / EMU || arrInchMargins[0]) + arrInchMargins[2])
		if (opts.debug) console.log('* Slide ' + arrObjSlides.length + ': emuSlideTabH (in) ........ = ' + (emuSlideTabH / EMU).toFixed(1))

		// C: Parse and store each cell's text into line array (**MAGIC HAPPENS HERE**)
		row.forEach((cell, iCell) => {
			// FIRST: REALITY-CHECK:
			if (!cell) cell = {}

			// DESIGN: Cells are henceforth {objects} with `text` and `opts`
			var lines: Array<string> = []

			// 1: Cleanse data
			if (!isNaN(cell) || typeof cell === 'string') {
				// Grab table formatting `opts` to use here so text style/format inherits as it should
				cell = { text: cell.toString(), opts: opts }
			} else if (typeof cell === 'object') {
				// ARG0: `text`
				if (typeof cell.text === 'number') cell.text = cell.text.toString()
				else if (typeof cell.text === 'undefined' || cell.text == null) cell.text = ''

				// ARG1: `options`
				var opt = cell.options || cell.opts || {}
				cell.opts = opt
			}
			// Capture some table options for use in other functions
			cell.opts.lineWeight = opts.lineWeight

			// 2: Create a cell object for each table column
			currRow.push({ text: '', opts: cell.opts })

			// 3: Parse cell contents into lines (**MAGIC HAPPENSS HERE**)
			var lines: Array<string> = parseTextToLines(cell, opts.colW[iCell] / ONEPT)
			arrCellsLines.push(lines)
			//if (opts.debug) console.log('Cell:'+iCell+' - lines:'+lines.length);

			// 4: Keep track of max line count within all row cells
			if (lines.length > intMaxLineCnt) {
				intMaxLineCnt = lines.length
				intMaxColIdx = iCell
			}
			var lineHeight = inch2Emu(((cell.opts.fontSize || opts.fontSize || DEF_FONT_SIZE) * LINEH_MODIFIER) / 100)
			// NOTE: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
			if (cell.opts && cell.opts.rowspan) lineHeight = 0

			// 5: Add cell margins to lineHeight (if any)
			if (cell.opts.margin) {
				if (cell.opts.margin[0]) lineHeight += (cell.opts.margin[0] * ONEPT) / intMaxLineCnt
				if (cell.opts.margin[2]) lineHeight += (cell.opts.margin[2] * ONEPT) / intMaxLineCnt
			}

			// Add to array
			arrCellsLineHeights.push(Math.round(lineHeight))
		})

		// D: AUTO-PAGING: Add text one-line-a-time to this row's cells until: lines are exhausted OR table H limit is hit
		for (var idx = 0; idx < intMaxLineCnt; idx++) {
			// 1: Add the current line to cell
			for (var col = 0; col < arrCellsLines.length; col++) {
				// A: Commit this slide to Presenation if table Height limit is hit
				if (emuTabCurrH + arrCellsLineHeights[intMaxColIdx] > emuSlideTabH) {
					if (opts.debug) console.log('--------------- New Slide Created ---------------')
					if (opts.debug)
						console.log(
							' (calc) ' + (emuTabCurrH / EMU).toFixed(1) + '+' + (arrCellsLineHeights[intMaxColIdx] / EMU).toFixed(1) + ' > ' + (emuSlideTabH / EMU).toFixed(1)
						)
					if (opts.debug) console.log('--------------- New Slide Created ---------------')
					// 1: Add the current row to table
					// NOTE: Edge cases can occur where we create a new slide only to have no more lines
					// ....: and then a blank row sits at the bottom of a table!
					// ....: Hence, we verify all cells have text before adding this final row.
					jQuery.each(currRow, (_idx, cell) => {
						if (cell.text.length > 0) {
							// IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
							arrRows.push(jQuery.extend(true, [], currRow))
							return false // break out of .each loop
						}
					})
					// 2: Add new Slide with current array of table rows
					arrObjSlides.push(jQuery.extend(true, [], arrRows))
					// 3: Empty rows for new Slide
					arrRows.length = 0
					// 4: Reset current table height for new Slide
					emuTabCurrH = 0 // This row's emuRowH w/b added below
					// 5: Empty current row's text (continue adding lines where we left off below)
					jQuery.each(currRow, (_idx, cell) => {
						cell.text = ''
					})
					// 6: Auto-Paging Options: addHeaderToEach
					if (opts.addHeaderToEach && arrObjTabHeadRows) arrRows = arrRows.concat(arrObjTabHeadRows)
				}

				// B: Add next line of text to this cell
				if (arrCellsLines[col][idx]) currRow[col].text += arrCellsLines[col][idx]
			}

			// 2: Add this new rows H to overall (use cell with the most lines as the determiner for overall row Height)
			emuTabCurrH += arrCellsLineHeights[intMaxColIdx]
		}

		if (opts.debug) console.log('-> ' + iRow + ' row done!')
		if (opts.debug) console.log('-> emuTabCurrH (in) . = ' + (emuTabCurrH / EMU).toFixed(1))

		// E: Flush row buffer - Add the current row to table, then truncate row cell array
		// IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
		if (currRow.length) arrRows.push(jQuery.extend(true, [], currRow))
		currRow.length = 0
	})

	// STEP 4-2: Flush final row buffer to slide
	arrObjSlides.push(jQuery.extend(true, [], arrRows))

	// LAST:
	if (opts.debug) {
		console.log('arrObjSlides count = ' + arrObjSlides.length)
		console.log(arrObjSlides)
	}
	return arrObjSlides
}
