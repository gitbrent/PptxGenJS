import { getUuid } from '../gen-utils'

import ElementInterface from './element-interface'

import Position from './position'
import RunProperties from './run-properties'

export default class CurrentDateElement implements ElementInterface {
    position
    runProperties
    fieldId
    dateFormat?: string

    constructor({ x, y, w, h, dateFormat, ...runOptions }, relations) {
        this.fieldId = getUuid('xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx')
        this.dateFormat =
            dateFormat && dateFormat > 0 && dateFormat < 14
                ? `datetime${dateFormat}`
                : 'datetime'
        this.position = new Position({
            x,
            y,
            w: w || 800000,
            h: h || 300000
        })
        this.runProperties = new RunProperties(runOptions, relations)
    }
    render(idx, presLayout, placeholder) {
        return `
		<p:sp>
		    <p:nvSpPr>
			    <p:cNvPr id="${idx + 1}" name="Datetime Placeholder ${idx + 1}"/>
			    <p:cNvSpPr txBox="1"></p:cNvSpPr>
			    <p:nvPr userDrawn="1" />
			</p:nvSpPr>

			<p:spPr>
			    ${this.position.render(presLayout)}
			    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:extLst><a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">
            <ma14:wrappingTextBoxFlag 
             val="0" 
             xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/>
          </a:ext></a:extLst>
			</p:spPr>

		    <p:txBody>
		        <a:bodyPr/>
		        <a:lstStyle><a:lvl1pPr>
		            ${this.runProperties.render('a:defRPr')}
		        </a:lvl1pPr></a:lstStyle>
                <a:p>
                    <a:fld id="{${this.fieldId}}" type="${this.dateFormat}">
                    <a:rPr /><a:t>‹today›</a:t>
                    </a:fld>
                    <a:endParaRPr />
                </a:p>
            </p:txBody>
        </p:sp>`
    }
}
