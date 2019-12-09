import { encodeXmlEntities } from '../gen-utils'

export default class TextFragment {
	text

	paragraphConfig
	runConfig

	constructor(text, paragraphConfig, runConfig) {
		this.text = text
		this.paragraphConfig = paragraphConfig
		this.runConfig = runConfig
	}

	render() {
		return `
		${this.paragraphConfig.render('a:pPr')}
        <a:r>
            ${this.runConfig.render('a:rPr')}
            <a:t>${encodeXmlEntities(this.text)}</a:t>
        </a:r>
        `
	}
}
