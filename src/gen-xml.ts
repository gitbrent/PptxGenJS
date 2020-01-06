import {
    CRLF,
    LAYOUT_IDX_SERIES_BASE,
    PLACEHOLDER_TYPES,
    SLDNUMFLDID,
    DEF_PRES_LAYOUT_NAME
} from './core-enums'
import { PowerPointShapes } from './core-shapes'
import {
    ILayout,
    ISlide,
    ISlideLayout,
    ISlideObject,
    ISlideRel,
    ISlideRelChart,
    ISlideRelMedia
} from './core-interfaces'
import { encodeXmlEntities, genXmlColorSelection } from './gen-utils'

import TextElement from './elements/text'
import ShapeElement from './elements/simple-shape'
import PlaceholderTextElement from './elements/placeholder-text'
import PlaceholderImageElement from './elements/placeholder-image'
import ImageElement from './elements/image'
import ChartElement from './elements/chart'
import SlideNumberElement from './elements/slide-number'
import TableElement from './elements/table'
import MediaElement from './elements/media'
import GroupElement from './elements/group'

/**
 * Transforms a slide or slideLayout to resulting XML string - Creates `ppt/slide*.xml`
 * @param {ISlide|ISlideLayout} slideObject - slide object created within createSlideObject
 * @return {string} XML string with <p:cSld> as the root
 */
function slideObjectToXml(slide: ISlide | ISlideLayout): string {
    let strSlideXml: string = slide.name
        ? `<p:cSld name="${slide.name}">`
        : '<p:cSld>'

    // STEP 1: Add background
    if (slide.bkgd) {
        strSlideXml += genXmlColorSelection(null, slide.bkgd)
    } else if (
        !slide.bkgd &&
        slide.name &&
        slide.name === DEF_PRES_LAYOUT_NAME
    ) {
        // NOTE: Default [white] background is needed on slideMaster1.xml
        // to avoid gray background in Keynote (and Finder previews)
        strSlideXml +=
            '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
    }

    // STEP 2: Add background image (using Strech) (if any)
    if (slide.bkgdImgRid) {
        // FIXME: We should be doing this in the slideLayout...
        strSlideXml +=
            '<p:bg>' +
            '<p:bgPr><a:blipFill dpi="0" rotWithShape="1">' +
            '<a:blip r:embed="rId' +
            slide.bkgdImgRid +
            '"><a:lum/></a:blip>' +
            '<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>' +
            '<a:effectLst/></p:bgPr>' +
            '</p:bg>'
    }

    // STEP 3: Continue slide by starting spTree node
    strSlideXml += '<p:spTree>'
    strSlideXml +=
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    strSlideXml +=
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    strSlideXml +=
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'

    // STEP 4: Loop over all Slide.data objects and add them to this slide
    slide.data.forEach((element: ISlideObject | TextElement, idx: number) => {
        if (element instanceof TextElement || element instanceof ImageElement) {
            const placeholder =
                slide['slideLayout'] &&
                slide['slideLayout'].getPlaceholder(element.placeholder)
            strSlideXml += element.render(idx, slide.presLayout, placeholder)
            return
        }
        if (
            element instanceof ShapeElement ||
            element instanceof PlaceholderTextElement ||
            element instanceof PlaceholderImageElement ||
            element instanceof ChartElement ||
            element instanceof TableElement ||
            element instanceof MediaElement ||
            element instanceof GroupElement ||
            element instanceof SlideNumberElement
        ) {
            strSlideXml += element.render(idx, slide.presLayout)
            return
        }
    })

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
 * @param {ISlide | ISlideLayout} slide - slide object whose relations are being transformed
 * @param {{ target: string; type: string }[]} defaultRels - array of default relations
 * @return {string} XML
 */
function slideObjectRelationsToXml(
    slide: ISlide | ISlideLayout,
    defaultRels: { target: string; type: string }[]
): string {
    let lastRid = 0 // stores maximum rId used for dynamic relations
    let strXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        CRLF +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'

    // STEP 1: Add all rels for this Slide
    slide.rels.forEach((rel: ISlideRel) => {
        lastRid = Math.max(lastRid, rel.rId)
        if (rel.type.toLowerCase().indexOf('hyperlink') > -1) {
            if (rel.data === 'slide') {
                strXml +=
                    '<Relationship Id="rId' +
                    rel.rId +
                    '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"' +
                    ' Target="slide' +
                    rel.Target +
                    '.xml"/>'
            } else {
                strXml +=
                    '<Relationship Id="rId' +
                    rel.rId +
                    '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"' +
                    ' Target="' +
                    rel.Target +
                    '" TargetMode="External"/>'
            }
        } else if (rel.type.toLowerCase().indexOf('notesSlide') > -1) {
            strXml +=
                '<Relationship Id="rId' +
                rel.rId +
                '" Target="' +
                rel.Target +
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>'
        }
    })
    ;(slide.relsChart || []).forEach((rel: ISlideRelChart) => {
        lastRid = Math.max(lastRid, rel.rId)
        strXml +=
            '<Relationship Id="rId' +
            rel.rId +
            '" Target="' +
            rel.Target +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>'
    })
    ;(slide.relsMedia || []).forEach((rel: ISlideRelMedia) => {
        lastRid = Math.max(lastRid, rel.rId)
        if (rel.type.toLowerCase().indexOf('image') > -1) {
            strXml +=
                '<Relationship Id="rId' +
                rel.rId +
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' +
                rel.Target +
                '"/>'
        } else if (rel.type.toLowerCase().indexOf('audio') > -1) {
            // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
            if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                strXml +=
                    '<Relationship Id="rId' +
                    rel.rId +
                    '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' +
                    rel.Target +
                    '"/>'
            else
                strXml +=
                    '<Relationship Id="rId' +
                    rel.rId +
                    '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="' +
                    rel.Target +
                    '"/>'
        } else if (rel.type.toLowerCase().indexOf('video') > -1) {
            // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
            if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                strXml +=
                    '<Relationship Id="rId' +
                    rel.rId +
                    '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' +
                    rel.Target +
                    '"/>'
            else
                strXml +=
                    '<Relationship Id="rId' +
                    rel.rId +
                    '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="' +
                    rel.Target +
                    '"/>'
        } else if (rel.type.toLowerCase().indexOf('online') > -1) {
            // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
            if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                strXml +=
                    '<Relationship Id="rId' +
                    rel.rId +
                    '" Type="http://schemas.microsoft.com/office/2007/relationships/image" Target="' +
                    rel.Target +
                    '"/>'
            else
                strXml +=
                    '<Relationship Id="rId' +
                    rel.rId +
                    '" Target="' +
                    rel.Target +
                    '" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>'
        }
    })

    // STEP 2: Add default rels
    defaultRels.forEach((rel, idx) => {
        strXml +=
            '<Relationship Id="rId' +
            (lastRid + idx + 1) +
            '" Type="' +
            rel.type +
            '" Target="' +
            rel.target +
            '"/>'
    })

    strXml += '</Relationships>'
    return strXml
}

/**
 * Generate an XML Placeholder
 * @param {ISlideObject} placeholderObj
 * @returns XML
 */
export function genXmlPlaceholder(placeholderObj: ISlideObject): string {
    if (!placeholderObj) return ''

    let placeholderIdx =
        placeholderObj.options && placeholderObj.options.placeholderIdx
            ? placeholderObj.options.placeholderIdx
            : ''
    let placeholderType =
        placeholderObj.options && placeholderObj.options.placeholderType
            ? placeholderObj.options.placeholderType
            : ''

    return `<p:ph
		${placeholderIdx ? ' idx="' + placeholderIdx + '"' : ''}
		${
            placeholderType && PLACEHOLDER_TYPES[placeholderType]
                ? ' type="' + PLACEHOLDER_TYPES[placeholderType] + '"'
                : ''
        }
		${
            placeholderObj.text && placeholderObj.text.length > 0
                ? ' hasCustomPrompt="1"'
                : ''
        }
		/>`
}

// XML-GEN: First 6 functions create the base /ppt files

/**
 * Generate XML ContentType
 * @param {ISlide[]} slides - slides
 * @param {ISlideLayout[]} slideLayouts - slide layouts
 * @param {ISlide} masterSlide - master slide
 * @returns XML
 */
export function makeXmlContTypes(
    slides: ISlide[],
    slideLayouts: ISlideLayout[],
    masterSlide?: ISlide
): string {
    let strXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
    strXml +=
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    strXml += '<Default Extension="xml" ContentType="application/xml"/>'
    strXml +=
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    strXml += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
    strXml += '<Default Extension="jpg" ContentType="image/jpg"/>'

    // STEP 1: Add standard/any media types used in Presenation
    strXml += '<Default Extension="png" ContentType="image/png"/>'
    strXml += '<Default Extension="gif" ContentType="image/gif"/>'
    strXml += '<Default Extension="m4v" ContentType="video/mp4"/>' // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)
    strXml += '<Default Extension="mp4" ContentType="video/mp4"/>' // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)
    slides.forEach(slide => {
        ;(slide.relsMedia || []).forEach(rel => {
            if (
                rel.type !== 'image' &&
                rel.type !== 'online' &&
                rel.type !== 'chart' &&
                rel.extn !== 'm4v' &&
                strXml.indexOf(rel.type) === -1
            ) {
                strXml +=
                    '<Default Extension="' +
                    rel.extn +
                    '" ContentType="' +
                    rel.type +
                    '"/>'
            }
        })
    })
    strXml +=
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
    strXml +=
        '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'

    // STEP 2: Add presentation and slide master(s)/slide(s)
    strXml +=
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
    strXml +=
        '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>'
    slides.forEach((slide, idx) => {
        strXml +=
            '<Override PartName="/ppt/slideMasters/slideMaster' +
            (idx + 1) +
            '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        strXml +=
            '<Override PartName="/ppt/slides/slide' +
            (idx + 1) +
            '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        // Add charts if any
        slide.relsChart.forEach(rel => {
            strXml +=
                ' <Override PartName="' +
                rel.Target +
                '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        })
    })

    // STEP 3: Core PPT
    strXml +=
        '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
    strXml +=
        '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>'
    strXml +=
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    strXml +=
        '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'

    // STEP 4: Add Slide Layouts
    slideLayouts.forEach((layout, idx) => {
        strXml +=
            '<Override PartName="/ppt/slideLayouts/slideLayout' +
            (idx + 1) +
            '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        ;(layout.relsChart || []).forEach(rel => {
            strXml +=
                ' <Override PartName="' +
                rel.Target +
                '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        })
    })

    // STEP 5: Add notes slide(s)
    slides.forEach((_slide, idx) => {
        strXml +=
            ' <Override PartName="/ppt/notesSlides/notesSlide' +
            (idx + 1) +
            '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
    })

    // STEP 6: Add rels
    masterSlide.relsChart.forEach(rel => {
        strXml +=
            ' <Override PartName="' +
            rel.Target +
            '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
    })
    masterSlide.relsMedia.forEach(rel => {
        if (
            rel.type !== 'image' &&
            rel.type !== 'online' &&
            rel.type !== 'chart' &&
            rel.extn !== 'm4v' &&
            strXml.indexOf(rel.type) === -1
        )
            strXml +=
                ' <Default Extension="' +
                rel.extn +
                '" ContentType="' +
                rel.type +
                '"/>'
    })

    // LAST: Finish XML (Resume core)
    strXml +=
        ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    strXml +=
        ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    strXml += '</Types>'

    return strXml
}

/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export function makeXmlRootRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
		<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
		</Relationships>`
}

/**
 * Creates `docProps/app.xml`
 * @param {ISlide[]} slides - Presenation Slides
 * @param {string} company - "Company" metadata
 * @returns XML
 */
export function makeXmlApp(slides: ISlide[], company: string): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
	<TotalTime>0</TotalTime>
	<Words>0</Words>
	<Application>Microsoft Office PowerPoint</Application>
	<PresentationFormat>On-screen Show (16:9)</PresentationFormat>
	<Paragraphs>0</Paragraphs>
	<Slides>${slides.length}</Slides>
	<Notes>${slides.length}</Notes>
	<HiddenSlides>0</HiddenSlides>
	<MMClips>0</MMClips>
	<ScaleCrop>false</ScaleCrop>
	<HeadingPairs>
		<vt:vector size="6" baseType="variant">
			<vt:variant><vt:lpstr>Fonts Used</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>2</vt:i4></vt:variant>
			<vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>1</vt:i4></vt:variant>
			<vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>${slides.length}</vt:i4></vt:variant>
		</vt:vector>
	</HeadingPairs>
	<TitlesOfParts>
		<vt:vector size="${slides.length + 1 + 2}" baseType="lpstr">
			<vt:lpstr>Arial</vt:lpstr>
			<vt:lpstr>Calibri</vt:lpstr>
			<vt:lpstr>Office Theme</vt:lpstr>
			${slides
                .map((_slideObj, idx) => {
                    return '<vt:lpstr>Slide ' + (idx + 1) + '</vt:lpstr>\n'
                })
                .join('')}
		</vt:vector>
	</TitlesOfParts>
	<Company>${company}</Company>
	<LinksUpToDate>false</LinksUpToDate>
	<SharedDoc>false</SharedDoc>
	<HyperlinksChanged>false</HyperlinksChanged>
	<AppVersion>16.0000</AppVersion>
	</Properties>`
}

/**
 * Creates `docProps/core.xml`
 * @param {string} title - metadata data
 * @param {string} company - metadata data
 * @param {string} author - metadata value
 * @param {string} revision - metadata value
 * @returns XML
 */
export function makeXmlCore(
    title: string,
    subject: string,
    author: string,
    revision: string
): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
		<dc:title>${encodeXmlEntities(title)}</dc:title>
		<dc:subject>${encodeXmlEntities(subject)}</dc:subject>
		<dc:creator>${encodeXmlEntities(author)}</dc:creator>
		<cp:lastModifiedBy>${encodeXmlEntities(author)}</cp:lastModifiedBy>
		<cp:revision>${revision}</cp:revision>
		<dcterms:created xsi:type="dcterms:W3CDTF">${new Date()
            .toISOString()
            .replace(/\.\d\d\dZ/, 'Z')}</dcterms:created>
		<dcterms:modified xsi:type="dcterms:W3CDTF">${new Date()
            .toISOString()
            .replace(/\.\d\d\dZ/, 'Z')}</dcterms:modified>
	</cp:coreProperties>`
}

/**
 * Creates `ppt/_rels/presentation.xml.rels`
 * @param {ISlide[]} slides - Presenation Slides
 * @returns XML
 */
export function makeXmlPresentationRels(slides: Array<ISlide>): string {
    let intRelNum = 1
    let strXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
    strXml +=
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    strXml +=
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
    for (let idx = 1; idx <= slides.length; idx++) {
        strXml +=
            '<Relationship Id="rId' +
            ++intRelNum +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' +
            idx +
            '.xml"/>'
    }
    intRelNum++
    strXml +=
        '<Relationship Id="rId' +
        intRelNum +
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>' +
        '<Relationship Id="rId' +
        (intRelNum + 1) +
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>' +
        '<Relationship Id="rId' +
        (intRelNum + 2) +
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>' +
        '<Relationship Id="rId' +
        (intRelNum + 3) +
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
        '<Relationship Id="rId' +
        (intRelNum + 4) +
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>' +
        '</Relationships>'

    return strXml
}

// XML-GEN: Functions that run 1-N times (once for each Slide)

/**
 * Generates XML for the slide file (`ppt/slides/slide1.xml`)
 * @param {ISlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export function makeXmlSlide(slide: ISlide): string {
    return (
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}` +
        `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
        `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"` +
        `${slide && slide.hidden ? ' show="0"' : ''}>` +
        `${slideObjectToXml(slide)}` +
        `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
    )
}

/**
 * Get text content of Notes from Slide
 * @param {ISlide} slide - the slide object to transform into XML
 * @return {string} notes text
 */
export function getNotesFromSlide(slide: ISlide): string {
    let notesText = ''

    slide.data.forEach(data => {
        if (data.type === 'notes') notesText += data.text
    })

    return notesText.replace(/\r*\n/g, CRLF)
}

/**
 * Generate XML for Notes Master (notesMaster1.xml)
 * @returns {string} XML
 */
export function makeXmlNotesMaster(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2971800" cy="458788"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="0"/><a:ext cx="2971800" cy="458788"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{5282F153-3F37-0F45-9E97-73ACFA13230C}" type="datetimeFigureOut"><a:rPr lang="en-US"/><a:t>7/23/19</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Image Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Notes Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US"/><a:t>Fifth level</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Footer Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="8685213"/><a:ext cx="2971800" cy="458787"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="7" name="Slide Number Placeholder 6"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="5"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="8685213"/><a:ext cx="2971800" cy="458787"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{CE5E9CC1-C706-0F49-92D6-E571CC5EEA8F}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991"/></p:ext></p:extLst></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:notesStyle></p:notesMaster>`
}

/**
 * Creates Notes Slide (`ppt/notesSlides/notesSlide1.xml`)
 * @param {ISlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export function makeXmlNotesSlide(slide: ISlide): string {
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        CRLF +
        '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>' +
        '<p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>' +
        '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>' +
        '</a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>' +
        '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/>' +
        '</p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' +
        '<p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
        '<a:rPr lang="en-US" dirty="0"/><a:t>' +
        encodeXmlEntities(getNotesFromSlide(slide)) +
        '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody>' +
        '</p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/>' +
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>' +
        '<p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr>' +
        '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>' +
        '<a:fld id="' +
        SLDNUMFLDID +
        '" type="slidenum">' +
        '<a:rPr lang="en-US"/><a:t>' +
        slide.number +
        '</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>' +
        '</p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">' +
        '<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991"/>' +
        '</p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
    )
}

/**
 * Generates the XML layout resource from a layout object
 * @param {ISlideLayout} layout - slide layout (master)
 * @return {string} XML
 */
export function makeXmlLayout(layout: ISlideLayout): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1">
		${slideObjectToXml(layout)}
		<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
}

/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param {ISlide} slide - slide object that represents master slide layout
 * @param {ISlideLayout[]} layouts - slide layouts
 * @return {string} XML
 */
export function makeXmlMaster(slide: ISlide, layouts: ISlideLayout[]): string {
    // NOTE: Pass layouts as static rels because they are not referenced any time
    let layoutDefs = layouts.map((_layoutDef, idx) => {
        return (
            '<p:sldLayoutId id="' +
            (LAYOUT_IDX_SERIES_BASE + idx) +
            '" r:id="rId' +
            (slide.rels.length + idx + 1) +
            '"/>'
        )
    })

    let strXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
    strXml +=
        '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    strXml += slideObjectToXml(slide)
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

    return strXml
}

/**
 * Generates XML string for a slide layout relation file
 * @param {number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @param {ISlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlSlideLayoutRel(
    layoutNumber: number,
    slideLayouts: ISlideLayout[]
): string {
    return slideObjectRelationsToXml(slideLayouts[layoutNumber - 1], [
        {
            target: '../slideMasters/slideMaster1.xml',
            type:
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
        }
    ])
}

/**
 * Creates `ppt/_rels/slide*.xml.rels`
 * @param {ISlide[]} slides
 * @param {ISlideLayout[]} slideLayouts - Slide Layout(s)
 * @param {number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export function makeXmlSlideRel(
    slides: ISlide[],
    slideLayouts: ISlideLayout[],
    slideNumber: number
): string {
    return slideObjectRelationsToXml(slides[slideNumber - 1], [
        {
            target:
                '../slideLayouts/slideLayout' +
                getLayoutIdxForSlide(slides, slideLayouts, slideNumber) +
                '.xml',
            type:
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
        },
        {
            target: '../notesSlides/notesSlide' + slideNumber + '.xml',
            type:
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
        }
    ])
}

/**
 * Generates XML string for a slide relation file.
 * @param {number} slideNumber - 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export function makeXmlNotesSlideRel(slideNumber: number): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${slideNumber}.xml"/>
		</Relationships>`
}

/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {ISlide} masterSlide - Slide object
 * @param {ISlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlMasterRel(
    masterSlide: ISlide,
    slideLayouts: ISlideLayout[]
): string {
    let defaultRels = slideLayouts.map((_layoutDef, idx) => {
        return {
            target: `../slideLayouts/slideLayout${idx + 1}.xml`,
            type:
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
        }
    })
    defaultRels.push({
        target: '../theme/theme1.xml',
        type:
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
    })

    return slideObjectRelationsToXml(masterSlide, defaultRels)
}

/**
 * Creates `ppt/notesMasters/_rels/notesMaster1.xml.rels`
 * @return {string} XML
 */
export function makeXmlNotesMasterRel(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
		</Relationships>`
}

/**
 * For the passed slide number, resolves name of a layout that is used for.
 * @param {ISlide[]} slides - srray of slides
 * @param {ISlideLayout[]} slideLayouts - array of slideLayouts
 * @param {number} slideNumber
 * @return {number} slide number
 */
function getLayoutIdxForSlide(
    slides: ISlide[],
    slideLayouts: ISlideLayout[],
    slideNumber: number
): number {
    for (let i = 0; i < slideLayouts.length; i++) {
        if (slideLayouts[i].name === slides[slideNumber - 1].slideLayout.name) {
            return i + 1
        }
    }

    // IMPORTANT: Return 1 (for `slideLayout1.xml`) when no def is found
    // So all objects are in Layout1 and every slide that references it uses this layout.
    return 1
}

// XML-GEN: Last 5 functions create root /ppt files

/**
 * Creates `ppt/theme/theme1.xml`
 * @return {string} XML
 */

/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {ISlide[]} slides - array of slides
 * @param {ILayout} pptLayout - presentation layout
 * @param {boolean} rtlMode - RTL mode
 * @return {string} XML
 */
export function makeXmlPresentation(
    slides: ISlide[],
    pptLayout: ILayout,
    rtlMode: boolean
): string {
    let strXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        CRLF +
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
        (rtlMode ? 'rtl="1" ' : '') +
        'saveSubsetFonts="1" autoCompressPictures="0">'

    // IMPORTANT: Steps 1-2-3 must be in this order or PPT will give corruption message on open!
    // STEP 1: Add slide master
    strXml +=
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'

    // STEP 2: Add all Slides
    strXml += '<p:sldIdLst>'
    for (let idx = 0; idx < slides.length; idx++) {
        strXml +=
            '<p:sldId id="' + (idx + 256) + '" r:id="rId' + (idx + 2) + '"/>'
    }
    strXml += '</p:sldIdLst>'

    // STEP 3: Add Notes Master (NOTE: length+2 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic))
    strXml +=
        '<p:notesMasterIdLst><p:notesMasterId r:id="rId' +
        (slides.length + 2) +
        '"/></p:notesMasterIdLst>'

    // STEP 4: Build SLIDE text styles
    strXml +=
        '<p:sldSz cx="' +
        pptLayout.width +
        '" cy="' +
        pptLayout.height +
        '"/>' +
        '<p:notesSz cx="' +
        pptLayout.height +
        '" cy="' +
        pptLayout.width +
        '"/>' +
        '<p:defaultTextStyle>' //+'<a:defPPr><a:defRPr lang="en-US"/></a:defPPr>'
    for (let idx = 1; idx < 10; idx++) {
        strXml +=
            '<a:lvl' +
            idx +
            'pPr marL="' +
            (idx - 1) * 457200 +
            '" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">' +
            '<a:defRPr sz="1800" kern="1200">' +
            '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
            '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>' +
            '</a:defRPr>' +
            '</a:lvl' +
            idx +
            'pPr>'
    }
    strXml += '</p:defaultTextStyle>'
    strXml += '</p:presentation>'

    return strXml
}

/**
 * Create `ppt/presProps.xml`
 * @return {string} XML
 */
export function makeXmlPresProps(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`
}

/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
export function makeXmlTableStyles(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`
}

/**
 * Creates `ppt/viewProps.xml`
 * @return {string} XML
 */
export function makeXmlViewProps(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1"><p:cViewPr varScale="1"><p:scale><a:sx n="136" d="100"/><a:sy n="136" d="100"/></p:scale><p:origin x="216" y="312"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`
}

export function getShapeInfo(shapeName) {
    if (!shapeName) return PowerPointShapes.RECTANGLE

    if (
        typeof shapeName === 'object' &&
        shapeName.name &&
        shapeName.displayName &&
        shapeName.avLst
    )
        return shapeName

    if (PowerPointShapes[shapeName]) return PowerPointShapes[shapeName]

    let objShape = Object.keys(PowerPointShapes).filter((key: string) => {
        return (
            PowerPointShapes[key].name === shapeName ||
            PowerPointShapes[key].displayName
        )
    })[0]
    if (typeof objShape !== 'undefined' && objShape !== null) return objShape

    return PowerPointShapes.RECTANGLE
}
