import {
    EMU,
    CHART_TYPES,
    ONEPT,
    PIECHART_COLORS,
    BARCHART_COLORS
} from './core-enums'
import { IShadowOptions } from './core-interfaces'

function correctGridLineOptions(glOpts) {
    if (!glOpts || glOpts.style === 'none') return
    if (
        glOpts.size !== undefined &&
        (isNaN(Number(glOpts.size)) || glOpts.size <= 0)
    ) {
        console.warn('Warning: chart.gridLine.size must be greater than 0.')
        delete glOpts.size // delete prop to used defaults
    }
    if (glOpts.style && ['solid', 'dash', 'dot'].indexOf(glOpts.style) < 0) {
        console.warn(
            'Warning: chart.gridLine.style options: `solid`, `dash`, `dot`.'
        )
        delete glOpts.style
    }

    return glOpts
}

/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {IShadowOptions} IShadowOptions - shadow options
 */
export function correctShadowOptions(IShadowOptions: IShadowOptions) {
    if (!IShadowOptions || IShadowOptions === null) return

    // OPT: `type`
    if (
        IShadowOptions.type !== 'outer' &&
        IShadowOptions.type !== 'inner' &&
        IShadowOptions.type !== 'none'
    ) {
        console.warn(
            'Warning: shadow.type options are `outer`, `inner` or `none`.'
        )
        IShadowOptions.type = 'outer'
    }

    // OPT: `angle`
    if (IShadowOptions.angle) {
        // A: REALITY-CHECK
        if (
            isNaN(Number(IShadowOptions.angle)) ||
            IShadowOptions.angle < 0 ||
            IShadowOptions.angle > 359
        ) {
            console.warn('Warning: shadow.angle can only be 0-359')
            IShadowOptions.angle = 270
        }

        // B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
        IShadowOptions.angle = Math.round(Number(IShadowOptions.angle))
    }

    // OPT: `opacity`
    if (IShadowOptions.opacity) {
        // A: REALITY-CHECK
        if (
            isNaN(Number(IShadowOptions.opacity)) ||
            IShadowOptions.opacity < 0 ||
            IShadowOptions.opacity > 1
        ) {
            console.warn('Warning: shadow.opacity can only be 0-1')
            IShadowOptions.opacity = 0.75
        }

        // B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
        IShadowOptions.opacity = Number(IShadowOptions.opacity)
    }
}

export const cleanChartOptions = options => {
    // STEP 1: TODO: check for reqd fields, correct type, etc
    // `type` exists in CHART_TYPES
    // Array.isArray(data)
    /*
		if ( Array.isArray(rel.data) && rel.data.length > 0 && typeof rel.data[0] === 'object'
			&& rel.data[0].labels && Array.isArray(rel.data[0].labels)
			&& rel.data[0].values && Array.isArray(rel.data[0].values) ) {
			obj = rel.data[0];
		}
		else {
			console.warn("USAGE: addChart( 'pie', [ {name:'Sales', labels:['Jan','Feb'], values:[10,20]} ], {x:1, y:1} )");
			return;
		}
		*/

    // B: Options: misc
    if (['bar', 'col'].indexOf(options.barDir || '') < 0) options.barDir = 'col'
    // IMPORTANT: 'bestFit' will cause issues with PPT-Online in some cases, so defualt to 'ctr'!
    if (
        [
            'bestFit',
            'b',
            'ctr',
            'inBase',
            'inEnd',
            'l',
            'outEnd',
            'r',
            't'
        ].indexOf(options.dataLabelPosition || '') < 0
    )
        options.dataLabelPosition =
            options.type === CHART_TYPES.PIE ||
            options.type === CHART_TYPES.DOUGHNUT
                ? 'bestFit'
                : 'ctr'
    options.dataLabelBkgrdColors =
        options.dataLabelBkgrdColors === true ||
        options.dataLabelBkgrdColors === false
            ? options.dataLabelBkgrdColors
            : false
    if (['b', 'l', 'r', 't', 'tr'].indexOf(options.legendPos || '') < 0)
        options.legendPos = 'r'
    // barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
    if (
        ['clustered', 'standard', 'stacked', 'percentStacked'].indexOf(
            options.barGrouping || ''
        ) < 0
    )
        options.barGrouping = 'standard'
    if (options.barGrouping.indexOf('tacked') > -1) {
        options.dataLabelPosition = 'ctr' // IMPORTANT: PPT-Online will not open Presentation when 'outEnd' etc is used on stacked!
        if (!options.barGapWidthPct) options.barGapWidthPct = 50
    }
    // 3D bar: ST_Shape
    if (
        [
            'cone',
            'coneToMax',
            'box',
            'cylinder',
            'pyramid',
            'pyramidToMax'
        ].indexOf(options.bar3DShape || '') < 0
    )
        options.bar3DShape = 'box'
    // lineDataSymbol: http://www.datypic.com/sc/ooxml/a-val-32.html
    // Spec has [plus,star,x] however neither PPT2013 nor PPT-Online support them
    if (
        [
            'circle',
            'dash',
            'diamond',
            'dot',
            'none',
            'square',
            'triangle'
        ].indexOf(options.lineDataSymbol || '') < 0
    )
        options.lineDataSymbol = 'circle'
    if (['gap', 'span'].indexOf(options.displayBlanksAs || '') < 0)
        options.displayBlanksAs = 'span'
    if (['standard', 'marker', 'filled'].indexOf(options.radarStyle || '') < 0)
        options.radarStyle = 'standard'
    options.lineDataSymbolSize =
        options.lineDataSymbolSize && !isNaN(options.lineDataSymbolSize)
            ? options.lineDataSymbolSize
            : 6
    options.lineDataSymbolLineSize =
        options.lineDataSymbolLineSize && !isNaN(options.lineDataSymbolLineSize)
            ? options.lineDataSymbolLineSize * ONEPT
            : 0.75 * ONEPT
    // `layout` allows the override of PPT defaults to maximize space
    if (options.layout) {
        ;['x', 'y', 'w', 'h'].forEach(key => {
            let val = options.layout[key]
            if (isNaN(Number(val)) || val < 0 || val > 1) {
                console.warn(
                    'Warning: chart.layout.' + key + ' can only be 0-1'
                )
                delete options.layout[key] // remove invalid value so that default will be used
            }
        })
    }

    // Set gridline defaults
    options.catGridLine =
        options.catGridLine ||
        (options.type === CHART_TYPES.SCATTER
            ? { color: 'D9D9D9', size: 1 }
            : { style: 'none' })
    options.valGridLine =
        options.valGridLine ||
        (options.type === CHART_TYPES.SCATTER
            ? { color: 'D9D9D9', size: 1 }
            : {})
    options.serGridLine =
        options.serGridLine ||
        (options.type === CHART_TYPES.SCATTER
            ? { color: 'D9D9D9', size: 1 }
            : { style: 'none' })
    correctGridLineOptions(options.catGridLine)
    correctGridLineOptions(options.valGridLine)
    correctGridLineOptions(options.serGridLine)
    correctShadowOptions(options.shadow)

    // C: Options: plotArea
    options.showDataTable =
        options.showDataTable === true || options.showDataTable === false
            ? options.showDataTable
            : false
    options.showDataTableHorzBorder =
        options.showDataTableHorzBorder === true ||
        options.showDataTableHorzBorder === false
            ? options.showDataTableHorzBorder
            : true
    options.showDataTableVertBorder =
        options.showDataTableVertBorder === true ||
        options.showDataTableVertBorder === false
            ? options.showDataTableVertBorder
            : true
    options.showDataTableOutline =
        options.showDataTableOutline === true ||
        options.showDataTableOutline === false
            ? options.showDataTableOutline
            : true
    options.showDataTableKeys =
        options.showDataTableKeys === true ||
        options.showDataTableKeys === false
            ? options.showDataTableKeys
            : true
    options.showLabel =
        options.showLabel === true || options.showLabel === false
            ? options.showLabel
            : false
    options.showLegend =
        options.showLegend === true || options.showLegend === false
            ? options.showLegend
            : false
    options.showPercent =
        options.showPercent === true || options.showPercent === false
            ? options.showPercent
            : true
    options.showTitle =
        options.showTitle === true || options.showTitle === false
            ? options.showTitle
            : false
    options.showValue =
        options.showValue === true || options.showValue === false
            ? options.showValue
            : false
    options.catAxisLineShow =
        typeof options.catAxisLineShow !== 'undefined'
            ? options.catAxisLineShow
            : true
    options.valAxisLineShow =
        typeof options.valAxisLineShow !== 'undefined'
            ? options.valAxisLineShow
            : true
    options.serAxisLineShow =
        typeof options.serAxisLineShow !== 'undefined'
            ? options.serAxisLineShow
            : true

    options.v3DRotX =
        !isNaN(options.v3DRotX) &&
        options.v3DRotX >= -90 &&
        options.v3DRotX <= 90
            ? options.v3DRotX
            : 30
    options.v3DRotY =
        !isNaN(options.v3DRotY) &&
        options.v3DRotY >= 0 &&
        options.v3DRotY <= 360
            ? options.v3DRotY
            : 30
    options.v3DRAngAx =
        options.v3DRAngAx === true || options.v3DRAngAx === false
            ? options.v3DRAngAx
            : true
    options.v3DPerspective =
        !isNaN(options.v3DPerspective) &&
        options.v3DPerspective >= 0 &&
        options.v3DPerspective <= 240
            ? options.v3DPerspective
            : 30

    // D: Options: chart
    options.barGapWidthPct =
        !isNaN(options.barGapWidthPct) &&
        options.barGapWidthPct >= 0 &&
        options.barGapWidthPct <= 1000
            ? options.barGapWidthPct
            : 150
    options.barGapDepthPct =
        !isNaN(options.barGapDepthPct) &&
        options.barGapDepthPct >= 0 &&
        options.barGapDepthPct <= 1000
            ? options.barGapDepthPct
            : 150

    options.chartColors = Array.isArray(options.chartColors)
        ? options.chartColors
        : options.type === CHART_TYPES.PIE ||
          options.type === CHART_TYPES.DOUGHNUT
        ? PIECHART_COLORS
        : BARCHART_COLORS
    options.chartColorsOpacity =
        options.chartColorsOpacity && !isNaN(options.chartColorsOpacity)
            ? options.chartColorsOpacity
            : null
    //
    options.border =
        options.border && typeof options.border === 'object'
            ? options.border
            : null
    if (options.border && (!options.border.pt || isNaN(options.border.pt)))
        options.border.pt = 1
    if (
        options.border &&
        (!options.border.color ||
            typeof options.border.color !== 'string' ||
            options.border.color.length !== 6)
    )
        options.border.color = '363636'
    //
    options.dataBorder =
        options.dataBorder && typeof options.dataBorder === 'object'
            ? options.dataBorder
            : null
    if (
        options.dataBorder &&
        (!options.dataBorder.pt || isNaN(options.dataBorder.pt))
    )
        options.dataBorder.pt = 0.75
    if (
        options.dataBorder &&
        (!options.dataBorder.color ||
            typeof options.dataBorder.color !== 'string' ||
            options.dataBorder.color.length !== 6)
    )
        options.dataBorder.color = 'F9F9F9'
    //
    if (!options.dataLabelFormatCode && options.type === CHART_TYPES.SCATTER)
        options.dataLabelFormatCode = 'General'
    options.dataLabelFormatCode =
        options.dataLabelFormatCode &&
        typeof options.dataLabelFormatCode === 'string'
            ? options.dataLabelFormatCode
            : options.type === CHART_TYPES.PIE ||
              options.type === CHART_TYPES.DOUGHNUT
            ? '0%'
            : '#,##0'
    //
    // Set default format for Scatter chart labels to custom string if not defined
    if (!options.dataLabelFormatScatter && options.type === CHART_TYPES.SCATTER)
        options.dataLabelFormatScatter = 'custom'
    //
    options.lineSize =
        typeof options.lineSize === 'number' ? options.lineSize : 2
    options.valAxisMajorUnit =
        typeof options.valAxisMajorUnit === 'number'
            ? options.valAxisMajorUnit
            : null
    options.valAxisCrossesAt = options.valAxisCrossesAt || 'autoZero'

    return options
}
