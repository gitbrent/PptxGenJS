import { CRLF } from '../core-enums'
import { encodeXmlEntities } from '../gen-utils'

import ParagraphProperties, {
    ParagraphPropertiesOptions
} from './paragraph-properties'
import RunProperties, { RunPropertiesOptions } from './run-properties'
import Relations from '../relations'

interface FragmentConfig {
    text: string
    options?: ParagraphPropertiesOptions & RunPropertiesOptions
}
export type FragmentOptions = string | number | FragmentConfig[]

export const buildFragments = (
    inputText: FragmentOptions,
    inputBreakLine: boolean,
    relations: Relations
) => {
    let fragments
    if (typeof inputText === 'string' || typeof inputText === 'number') {
        fragments = [{ text: inputText.toString(), options: {} }]
    } else {
        fragments = inputText
    }
    if (!fragments) return []

    return fragments.flatMap(({ text: fragmentText, options }, idx) => {
        let text = fragmentText.replace(/\r*\n/g, CRLF)
        let breakLine = inputBreakLine || options.breakLine || false

        if (text.indexOf(CRLF) > -1) {
            // Remove trailing linebreak (if any) so the "if" below doesnt create a double CRLF+CRLF line ending!
            text = text.replace(/\r\n$/g, '')
            // Plain strings like "hello \n world" or "first line\n" need to have lineBreaks set to become 2 separate lines as intended
            breakLine = true
        }

        const paragraphProperties = new ParagraphProperties({
            bullet: options.bullet,
            align: options.align,
            rtlMode: options.rtlMode,
            lineSpacing: options.lineSpacing,
            indentLevel: options.indentLevel,
            paraSpaceBefore: options.paraSpaceBefore,
            paraSpaceAfter: options.paraSpaceAfter
        })
        const runProperties = new RunProperties(
            {
                lang: options.lang,
                fontFace: options.fontFace,
                fontSize: options.fontSize,
                charSpacing: options.charSpacing,
                color: options.color,
                bold: options.bold,
                italic: options.italic,
                strike: options.strike,
                underline: options.underline,
                subscript: options.subscript,
                superscript: options.superscript,
                outline: options.outline,
                hyperlink: options.hyperlink
            },
            relations
        )

        if (breakLine) {
            return text.split(CRLF).map((line, lineIdx) => {
                return new TextFragment(
                    line,
                    paragraphProperties,
                    runProperties
                )
            })
        } else {
            return new TextFragment(text, paragraphProperties, runProperties)
        }
    })
}

export default class TextFragment {
    text: string

    paragraphConfig: ParagraphProperties
    runConfig: RunProperties

    constructor(
        text: string,
        paragraphConfig: ParagraphProperties,
        runConfig: RunProperties
    ) {
        this.text = text
        this.paragraphConfig = paragraphConfig
        this.runConfig = runConfig
    }

    render(presLayout) {
        return `
		${this.paragraphConfig.render(presLayout, 'a:pPr')}
        <a:r>
            ${this.runConfig.render('a:rPr')}
            <a:t>${encodeXmlEntities(this.text)}</a:t>
        </a:r>
        `
    }
}
