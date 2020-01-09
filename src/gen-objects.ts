/**
 * PptxGenJS: Slide object generators
 */

import { ONEPT, SLIDE_OBJECT_TYPES } from './core-enums'
import { ITableCell, IText } from './core-interfaces'
import { getSlidesForTableRows } from './gen-tables'
import { encodeXmlEntities } from './gen-utils'

import Slide from './slide'
import { Master } from './slideLayouts'

import TextElement from './elements/text'
import ShapeElement from './elements/simple-shape'
import PlaceholderTextElement from './elements/placeholder-text'
import ImageElement from './elements/image'
import ChartElement from './elements/chart'
import SlideNumberElement from './elements/slide-number'

/**
 * Adds placeholder objects to slide
 * @param {Slide} slide - slide object containing layouts
 */
export function addPlaceholdersToSlideLayouts(slide: Slide) {
    // Add all placeholders on this Slide that dont already exist
    ;(slide.slideLayout.data || []).forEach(slideLayoutObj => {
        if (slideLayoutObj instanceof PlaceholderTextElement) {
            // A: Search for this placeholder on Slide before we add
            // NOTE: Check to ensure a placeholder does not already exist on the Slide
            // They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
            if (
                slide.data.filter(slideObj => {
                    return slideObj.placeholder === slideLayoutObj.name
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
 * Parses text/text-objects from `addText()` and `addTable()` methods; creates 'hyperlink'-type Slide Rels for each hyperlink found
 * @param {Slide} target - slide object that any hyperlinks will be be added to
 * @param {number | string | IText | IText[] | ITableCell[][]} text - text to parse
 */
function createHyperlinkRels(
    target: Slide,
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
