import { encodeXmlEntities } from '../gen-utils'

import ParagraphProperties from './paragraph-properties'
import RunProperties from './run-properties'

export default class TextFragment {
    text

    paragraphConfig: ParagraphProperties
    runConfig: RunProperties

    constructor(text, paragraphConfig, runConfig) {
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
