import { getUuid } from '../gen-utils'

import ElementInterface from './element-interface'

import Position from './position'
import RunProperties from './run-properties'

export default class SlideNumberElement implements ElementInterface {
    position
    runProperties
    fieldId

    constructor({ x, y, w, h, ...runOptions }) {
        this.fieldId = getUuid('xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx')
        this.position = new Position({
            x,
            y,
            w: w || 800000,
            h: h || 300000
        })
        this.runProperties = new RunProperties(runOptions)
    }
    render(idx, presLayout, placeholder) {
        return `
		<p:sp>
		    <p:nvSpPr>
			    <p:cNvPr id="${idx + 1}" name="Slide Number Placeholder 24"/>
			    <p:cNvSpPr txBox="1"></p:cNvSpPr>
			    <p:nvPr userDrawn="1" />
			</p:nvSpPr>

			<p:spPr>
			    ${this.position.render(presLayout)}
			    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:extLst><a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">
                    <ma14:wrappingTextBoxFlag val="0"
                        xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/>
                </a:ext></a:extLst>
			</p:spPr>

		    <p:txBody>
		        <a:bodyPr/>
		        <a:lstStyle><a:lvl1pPr>
		            ${this.runProperties.render('a:defRPr')}
		        </a:lvl1pPr></a:lstStyle>
                <a:p>
                    <a:fld id="{${this.fieldId}}" type="slidenum">
                    <a:rPr lang="en-US"/><a:t>‹N°›</a:t>
                    </a:fld>
                    <a:endParaRPr lang="en-US"/>
                </a:p>
            </p:txBody>
        </p:sp>`
    }
}
