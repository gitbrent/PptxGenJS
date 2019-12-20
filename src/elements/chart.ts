import { SLIDE_OBJECT_TYPES } from '../core-enums'

import { cleanChartOptions } from '../chart-options'
import Position from './position'

let _chartCounter: number = 0

/*
 * This class manages the inclusing of a chart in a slide
 * The logic to generate the chart itself (and then its ooxml is in a separate
 * class)
 */

export default class ChartElement {
	type = SLIDE_OBJECT_TYPES.newtext
	chartId
	position

	constructor(type, data, opts, relations) {
		// DESIGN: `type` can an object (ex: `pptx.charts.DOUGHNUT`) or an array of chart objects
		// EX: addChartDefinition([ { type:pptx.charts.BAR, data:{name:'', labels:[], values[]} }, {<etc>} ])
		// Multi-Type Charts
		let tmpOpt
		let tmpData = []

		if (Array.isArray(type)) {
			// For multi-type charts there needs to be data for each type,
			// as well as a single data source for non-series operations.
			// The data is indexed below to keep the data in order when segmented
			// into types.
			type.forEach(obj => {
				tmpData = tmpData.concat(obj.data)
			})
			tmpOpt = data || opts
		} else {
			tmpData = data
			tmpOpt = opts
		}
		tmpData.forEach((item, i) => {
			item.index = i
		})
		const options = tmpOpt && typeof tmpOpt === 'object' ? tmpOpt : {}

		this.position = new Position({
			x: typeof options.x !== 'undefined' && options.x != null ? options.x : 1,
			y: typeof options.y !== 'undefined' && options.y != null ? options.y : 1,
			w: options.w || '50%',
			h: options.w || '50%',
		})

		// This should probably be managed somewhere else (within register). We
		// keep it that way as long as masters work differently than slides.
		let globalChartId = ++_chartCounter

		options.type = type
		this.chartId = relations.registerChart(globalChartId, cleanChartOptions(options), tmpData)
	}

	render(idx, presLayout) {
		return `
        <p:graphicFrame>
		    <p:nvGraphicFramePr>
			    <p:cNvPr id="${idx + 2}" name="Chart ${idx + 1}"/>
			    <p:cNvGraphicFramePr/>
			    <p:nvPr></p:nvPr>
			</p:nvGraphicFramePr>
			${this.position.render(presLayout, 'p:xfrm')}
			<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
				<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                    <c:chart 
                        r:id="rId${this.chartId}" 
                        xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>
                </a:graphicData>
			</a:graphic>
		</p:graphicFrame>`
	}
}
