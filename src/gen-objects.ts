/**
 * PptxGenJS: Slide object generators
 */

import { ONEPT, SLIDE_OBJECT_TYPES } from './core-enums'
import { ISlide, ISlideLayout, ITableCell, IText } from './core-interfaces'
import { getSlidesForTableRows } from './gen-tables'
import { encodeXmlEntities } from './gen-utils'

import TextElement from './elements/text'
import ShapeElement from './elements/simple-shape'
import PlaceholderTextElement from './elements/placeholder-text'
import ImageElement from './elements/image'
import ChartElement from './elements/chart'
import SlideNumberElement from './elements/slide-number'

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
        text: notes
    })
}

/**
 * Adds placeholder objects to slide
 * @param {ISlide} slide - slide object containing layouts
 */
export function addPlaceholdersToSlideLayouts(slide: ISlide) {
    // Add all placeholders on this Slide that dont already exist
    ;(slide.slideLayout.data || []).forEach(slideLayoutObj => {
        if (slideLayoutObj instanceof PlaceholderTextElement) {
            // A: Search for this placeholder on Slide before we add
            // NOTE: Check to ensure a placeholder does not already exist on the Slide
            // They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
            if (
                slide.data.filter(slideObj => {
                    const placeholder =
                        slideObj.placeholder ||
                        (slideObj.options && slideObj.options.placeholder)
                    return placeholder === slideLayoutObj.name
                }).length === 0
            ) {
                if (slideLayoutObj.placeholderType !== 'pic') {
                    slide.data.push(
                        new TextElement(
                            '',
                            { placeholder: slideLayoutObj.name },
                            () => null
                        )
                    )
                }
            }
        }
    })
}

/* -------------------------------------------------------------------------------- */

/**
 * Adds a background image or color to a slide definition.
 * @param {String|Object} bkg - color string or an object with image definition
 * @param {ISlide} target - slide object that the background is set to
 */
function addBackgroundDefinition(
    bkg: string | { src?: string; path?: string; data?: string },
    target: ISlide | ISlideLayout
) {
    if (typeof bkg === 'object' && (bkg.src || bkg.path || bkg.data)) {
        // Allow the use of only the data key (`path` isnt reqd)
        bkg.src = bkg.src || bkg.path || null
        if (!bkg.src) bkg.src = 'preencoded.png'
        let strImgExtn = (bkg.src.split('.').pop() || 'png').split('?')[0] // Handle "blah.jpg?width=540" etc.
        if (strImgExtn === 'jpg') strImgExtn = 'jpeg' // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup

        let intRels = target.relsMedia.length + 1
        target.relsMedia.push({
            path: bkg.src,
            type: SLIDE_OBJECT_TYPES.image,
            extn: strImgExtn,
            data: bkg.data || null,
            rId: intRels,
            Target:
                '../media/image' +
                (target.relsMedia.length + 1) +
                '.' +
                strImgExtn
        })
        target.bkgdImgRid = intRels
    } else if (bkg && typeof bkg === 'string') {
        target.bkgd = bkg
    }
}

/**
 * Parses text/text-objects from `addText()` and `addTable()` methods; creates 'hyperlink'-type Slide Rels for each hyperlink found
 * @param {ISlide} target - slide object that any hyperlinks will be be added to
 * @param {number | string | IText | IText[] | ITableCell[][]} text - text to parse
 */
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
