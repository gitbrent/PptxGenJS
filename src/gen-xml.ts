/**
* PptxGenJS: XML Generation
*/

import {
	CRLF, ONEPT,
	REGEX_HEX_COLOR,
	BARCHART_COLORS,
	DEF_CHART_GRIDLINE,
	DEF_SHAPE_SHADOW,
	DEF_TEXT_SHADOW,
	CHART_TYPES,
	LETTERS,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY, AXIS_ID_SERIES_PRIMARY, BULLET_TYPES, DEF_FONT_TITLE_SIZE, DEF_FONT_COLOR, DEF_FONT_SIZE,
	LAYOUT_IDX_SERIES_BASE, PLACEHOLDER_TYPES
} from './enums';
import { gObjPptx } from './pptxgen'
import { getMix, getUuid, encodeXmlEntities } from './utils';

// TODO: export default class GenXml - encapsulate these func - we onyl nee dot expose 1-2

/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 */
export function makeXmlCharts(rel) {
	// HELPER FUNCS:
	function hasArea(chartType: CHART_TYPES) {
		function has(type) {
			return chartType.some(function(item) {
				return item.type.name === type;
			});
		}
		if (Array.isArray(chartType)) {
			return has('area');
		}
		return chartType === 'area';
	}

	/* ----------------------------------------------------------------------- */

	// STEP 1: Create chart
	{
		var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
		// CHARTSPACE: BEGIN vvv
		strXml += '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		strXml += '<c:date1904 val="0"/>';  // ppt defaults to 1904 dates, excel to 1900
		strXml += '<c:chart>';

		// OPTION: Title
		if (rel.opts.showTitle) {
			strXml += genXmlTitle({
				title: rel.opts.title || 'Chart Title',
				fontSize: rel.opts.titleFontSize || DEF_FONT_TITLE_SIZE,
				color: rel.opts.titleColor,
				fontFace: rel.opts.titleFontFace,
				rotate: rel.opts.titleRotate,
				titleAlign: rel.opts.titleAlign,
				titlePos: rel.opts.titlePos
			});
			strXml += '<c:autoTitleDeleted val="0"/>';
		}
		else {
			// NOTE: Add autoTitleDeleted tag in else to prevent default creation of chart title even when showTitle is set to false
			strXml += '<c:autoTitleDeleted val="1"/>';
		}
		// Add 3D view tag
		if (rel.opts.type.name == 'bar3D') {
			strXml += '<c:view3D>';
			strXml += ' <c:rotX val="' + rel.opts.v3DRotX + '"/>';
			strXml += ' <c:rotY val="' + rel.opts.v3DRotY + '"/>';
			strXml += ' <c:rAngAx val="' + (rel.opts.v3DRAngAx == false ? 0 : 1) + '"/>';
			strXml += ' <c:perspective val="' + rel.opts.v3DPerspective + '"/>';
			strXml += '</c:view3D>';
		}

		strXml += '<c:plotArea>';
		// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
		if (rel.opts.layout) {
			strXml += '<c:layout>';
			strXml += ' <c:manualLayout>';
			strXml += '  <c:layoutTarget val="inner" />';
			strXml += '  <c:xMode val="edge" />';
			strXml += '  <c:yMode val="edge" />';
			strXml += '  <c:x val="' + (rel.opts.layout.x || 0) + '" />';
			strXml += '  <c:y val="' + (rel.opts.layout.y || 0) + '" />';
			strXml += '  <c:w val="' + (rel.opts.layout.w || 1) + '" />';
			strXml += '  <c:h val="' + (rel.opts.layout.h || 1) + '" />';
			strXml += ' </c:manualLayout>';
			strXml += '</c:layout>';
		}
		else {
			strXml += '<c:layout/>';
		}
	}

	var usesSecondaryValAxis = false;

	// A: Create Chart XML -----------------------------------------------------------
	if (Array.isArray(rel.opts.type)) {
		rel.opts.type.forEach((type) => {
			var chartType = type.type.name;
			var data = type.data;
			var options = getMix(rel.opts, type.options);
			var valAxisId = options['secondaryValAxis'] ? AXIS_ID_VALUE_SECONDARY : AXIS_ID_VALUE_PRIMARY;
			var catAxisId = options['secondaryCatAxis'] ? AXIS_ID_CATEGORY_SECONDARY : AXIS_ID_CATEGORY_PRIMARY;
			var isMultiTypeChart = true;
			usesSecondaryValAxis = usesSecondaryValAxis || options['secondaryValAxis'];
			strXml += makeChartType(chartType, data, options, valAxisId, catAxisId, isMultiTypeChart);
		});
	}
	else {
		var chartType = rel.opts.type.name;
		var isMultiTypeChart = false;
		strXml += makeChartType(chartType, rel.data, rel.opts, AXIS_ID_VALUE_PRIMARY, AXIS_ID_CATEGORY_PRIMARY, isMultiTypeChart);
	}

	// B: Axes -----------------------------------------------------------
	if (rel.opts.type.name !== 'pie' && rel.opts.type.name !== 'doughnut') {
		// Param check
		if (rel.opts.valAxes && !usesSecondaryValAxis) {
			throw new Error('Secondary axis must be used by one of the multiple charts');
		}

		if (rel.opts.catAxes) {
			if (!rel.opts.valAxes || rel.opts.valAxes.length !== rel.opts.catAxes.length) {
				throw new Error('There must be the same number of value and category axes.');
			}
			strXml += makeCatAxis(getMix(rel.opts, rel.opts.catAxes[0]), AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY);
			if (rel.opts.catAxes[1]) {
				strXml += makeCatAxis(getMix(rel.opts, rel.opts.catAxes[1]), AXIS_ID_CATEGORY_SECONDARY, AXIS_ID_VALUE_PRIMARY);
			}
		}
		else {
			strXml += makeCatAxis(rel.opts, AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY);
		}

		rel.opts.hasArea = hasArea(rel.opts.type);

		if (rel.opts.valAxes) {
			strXml += makeValueAxis(getMix(rel.opts, rel.opts.valAxes[0]), AXIS_ID_VALUE_PRIMARY);
			if (rel.opts.valAxes[1]) {
				strXml += makeValueAxis(getMix(rel.opts, rel.opts.valAxes[1]), AXIS_ID_VALUE_SECONDARY);
			}
		}
		else {
			strXml += makeValueAxis(rel.opts, AXIS_ID_VALUE_PRIMARY);

			// Add series axis for 3D bar
			if (rel.opts.type.name == 'bar3D') {
				strXml += makeSerAxis(rel.opts, AXIS_ID_SERIES_PRIMARY, AXIS_ID_VALUE_PRIMARY)
			}
		}
	}

	// C: Chart Properties and plotArea Options: Border, Data Table, Fill, Legend
	{
		// NOTE: DataTable goes between '</c:valAx>' and '<c:spPr>'
		if (rel.opts.showDataTable) {
			strXml += '<c:dTable>';
			strXml += '  <c:showHorzBorder val="' + (rel.opts.showDataTableHorzBorder == false ? 0 : 1) + '"/>';
			strXml += '  <c:showVertBorder val="' + (rel.opts.showDataTableVertBorder == false ? 0 : 1) + '"/>';
			strXml += '  <c:showOutline    val="' + (rel.opts.showDataTableOutline == false ? 0 : 1) + '"/>';
			strXml += '  <c:showKeys       val="' + (rel.opts.showDataTableKeys == false ? 0 : 1) + '"/>';
			strXml += '  <c:spPr>';
			strXml += '    <a:noFill/>';
			strXml += '    <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>';
			strXml += '    <a:effectLst/>';
			strXml += '  </c:spPr>';
			strXml += '  <c:txPr>\
						  <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
						  <a:lstStyle/>\
						  <a:p>\
							<a:pPr rtl="0">\
							  <a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
								<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>\
								<a:latin typeface="+mn-lt"/>\
								<a:ea typeface="+mn-ea"/>\
								<a:cs typeface="+mn-cs"/>\
							  </a:defRPr>\
							</a:pPr>\
							<a:endParaRPr lang="en-US"/>\
						  </a:p>\
						</c:txPr>\
					  </c:dTable>';
		}

		strXml += '  <c:spPr>';

		// OPTION: Fill
		strXml += (rel.opts.fill ? genXmlColorSelection(rel.opts.fill) : '<a:noFill/>');

		// OPTION: Border
		strXml += (rel.opts.border ? '<a:ln w="' + (rel.opts.border.pt * ONEPT) + '"' + ' cap="flat">' + genXmlColorSelection(rel.opts.border.color) + '</a:ln>' : '<a:ln><a:noFill/></a:ln>');

		// Close shapeProp/plotArea before Legend
		strXml += '    <a:effectLst/>';
		strXml += '  </c:spPr>';
		strXml += '</c:plotArea>';

		// OPTION: Legend
		// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
		if (rel.opts.showLegend) {
			strXml += '<c:legend>';
			strXml += '<c:legendPos val="' + rel.opts.legendPos + '"/>';
			strXml += '<c:layout/>';
			strXml += '<c:overlay val="0"/>';
			if (rel.opts.legendFontFace || rel.opts.legendFontSize || rel.opts.legendColor) {
				strXml += '<c:txPr>';
				strXml += '  <a:bodyPr/>';
				strXml += '  <a:lstStyle/>';
				strXml += '  <a:p>';
				strXml += '    <a:pPr>';
				strXml += (rel.opts.legendFontSize ? '<a:defRPr sz="' + (Number(rel.opts.legendFontSize) * 100) + '">' : '<a:defRPr>');
				if (rel.opts.legendColor) strXml += genXmlColorSelection(rel.opts.legendColor);
				if (rel.opts.legendFontFace) strXml += '<a:latin typeface="' + rel.opts.legendFontFace + '"/>';
				if (rel.opts.legendFontFace) strXml += '<a:cs    typeface="' + rel.opts.legendFontFace + '"/>';
				strXml += '      </a:defRPr>';
				strXml += '    </a:pPr>';
				strXml += '    <a:endParaRPr lang="en-US"/>';
				strXml += '  </a:p>';
				strXml += '</c:txPr>';
			}
			strXml += '</c:legend>';
		}
	}

	strXml += '  <c:plotVisOnly val="1"/>';
	strXml += '  <c:dispBlanksAs val="' + rel.opts.displayBlanksAs + '"/>';
	if (rel.opts.type.name === 'scatter') strXml += '<c:showDLblsOverMax val="1"/>';

	strXml += '</c:chart>';

	// D: CHARTSPACE SHAPE PROPS
	strXml += '<c:spPr>';
	strXml += '  <a:noFill/>';
	strXml += '  <a:ln w="12700" cap="flat"><a:noFill/><a:miter lim="400000"/></a:ln>';
	strXml += '  <a:effectLst/>';
	strXml += '</c:spPr>';

	// E: DATA (Add relID)
	strXml += '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>';

	// LAST: chartSpace end
	strXml += '</c:chartSpace>';

	return strXml;
}

/**
 * Create XML string for any given chart type
 * @example: <c:bubbleChart> or <c:lineChart>
 *
 * @param {String} CHART_TYPES key
 * @param {String} data
 * @param {object} opts
 * @param {String} valAxisId
 * @param {String} catAxisId
 * @param {boolean} isMultiTypeChart
 */
function makeChartType(
	chartType: CHART_TYPES,
	data: Array<{ index: number, name: string }>,
	opts: { barDir: string, barGrouping: string, chartColors: Array<string>, chartColorsOpacity: number, radarStyle: string },
	valAxisId: string, catAxisId: string, isMultiTypeChart: boolean) {
	// NOTE: "Chart Range" (as shown in "select Chart Area dialog") is calculated.
	// ....: Ensure each X/Y Axis/Col has same row height (esp. applicable to XY Scatter where X can often be larger than Y's)
	var strXml = '';

	switch (chartType) {
		case 'area':
		case 'bar':
		case 'bar3D':
		case 'line':
		case 'radar':
			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>';
			if (chartType == 'bar' || chartType == 'bar3D') {
				strXml += '<c:barDir val="' + opts.barDir + '"/>';
				strXml += '<c:grouping val="' + opts.barGrouping + '"/>';
			}

			if (chartType == 'radar') {
				strXml += '<c:radarStyle val="' + opts.radarStyle + '"/>';
			}

			strXml += '<c:varyColors val="0"/>';

			// 2: "Series" block for every data row
			/* EX:
				data: [
				 {
				   name: 'Region 1',
				   labels: ['April', 'May', 'June', 'July'],
				   values: [17, 26, 53, 96]
				 },
				 {
				   name: 'Region 2',
				   labels: ['April', 'May', 'June', 'July'],
				   values: [55, 43, 70, 58]
				 }
				]
			*/
			var colorIndex = -1; // Maintain the color index by region
			data.forEach((obj) => {
				colorIndex++;
				var idx = obj.index;
				strXml += '<c:ser>';
				strXml += '  <c:idx val="' + idx + '"/>';
				strXml += '  <c:order val="' + idx + '"/>';
				strXml += '  <c:tx>';
				strXml += '    <c:strRef>';
				strXml += '      <c:f>Sheet1!$' + getExcelColName(idx + 1) + '$1</c:f>';
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt></c:strCache>';
				strXml += '    </c:strRef>';
				strXml += '  </c:tx>';
				strXml += '  <c:invertIfNegative val="0"/>';

				// Fill and Border
				var strSerColor = opts.chartColors[colorIndex % opts.chartColors.length];

				strXml += '  <c:spPr>';
				if (strSerColor == 'transparent') {
					strXml += '<a:noFill/>';
				}
				else if (opts.chartColorsOpacity) {
					strXml += '<a:solidFill>' + createColorElement(strSerColor, '<a:alpha val="' + opts.chartColorsOpacity + '000"/>') + '</a:solidFill>';
				}
				else {
					strXml += '<a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
				}

				if (chartType == 'line') {
					if (opts.lineSize == 0) {
						strXml += '<a:ln><a:noFill/></a:ln>';
					}
					else {
						strXml += '<a:ln w="' + (opts.lineSize * ONEPT) + '" cap="flat"><a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
						strXml += '<a:prstDash val="' + (opts.lineDash || "solid") + '"/><a:round/></a:ln>';
					}
				}
				else if (opts.dataBorder) {
					strXml += '<a:ln w="' + (opts.dataBorder.pt * ONEPT) + '" cap="flat"><a:solidFill>' + createColorElement(opts.dataBorder.color) + '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
				}

				strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);

				strXml += '  </c:spPr>';

				// Data Labels per series
				// [20190117] NOTE: Adding these to RADAR chart causes unrecoverable corruption!
				if (chartType != 'radar') {
					strXml += '  <c:dLbls>';
					strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
					if (opts.dataLabelBkgrdColors) {
						strXml += '    <c:spPr>';
						strXml += '       <a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
						strXml += '    </c:spPr>';
					}
					strXml += '    <c:txPr>';
					strXml += '      <a:bodyPr/>';
					strXml += '      <a:lstStyle/>';
					strXml += '      <a:p><a:pPr>';
					strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
					strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
					strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
					strXml += '        </a:defRPr>';
					strXml += '      </a:pPr></a:p>';
					strXml += '    </c:txPr>';
					// Setting dLblPos tag for bar3D seems to break the generated chart
					if (chartType != 'area' && chartType != 'bar3D') {
						strXml += '<c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>';
					}
					strXml += '    <c:showLegendKey val="0"/>';
					strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
					strXml += '    <c:showCatName val="0"/>';
					strXml += '    <c:showSerName val="0"/>';
					strXml += '    <c:showPercent val="0"/>';
					strXml += '    <c:showBubbleSize val="0"/>';
					strXml += '    <c:showLeaderLines val="0"/>';
					strXml += '  </c:dLbls>';
				}

				// 'c:marker' tag: `lineDataSymbol`
				if (chartType == 'line' || chartType == 'radar') {
					strXml += '<c:marker>';
					strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>';
					if (opts.lineDataSymbolSize) {
						// Defaults to "auto" otherwise (but this is usually too small, so there is a default)
						strXml += '  <c:size val="' + opts.lineDataSymbolSize + '"/>';
					}
					strXml += '  <c:spPr>';
					strXml += '    <a:solidFill>' + createColorElement(opts.chartColors[(idx + 1 > opts.chartColors.length ? (Math.floor(Math.random() * opts.chartColors.length)) : idx)]) + '</a:solidFill>';

					var symbolLineColor = opts.lineDataSymbolLineColor || strSerColor;
					strXml += '    <a:ln w="' + opts.lineDataSymbolLineSize + '" cap="flat"><a:solidFill>' + createColorElement(symbolLineColor) + '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
					strXml += '    <a:effectLst/>';
					strXml += '  </c:spPr>';
					strXml += '</c:marker>';
				}

				// Color chart bars various colors
				// Allow users with a single data set to pass their own array of colors (check for this using != ours)
				if ((chartType == 'bar' || chartType == 'bar3D') && (data.length === 1 || opts.valueBarColors) && opts.chartColors != BARCHART_COLORS) {
					// Series Data Point colors
					obj.values.forEach(function(value, index) {
						var arrColors = (value < 0 ? (opts.invertedColors || BARCHART_COLORS) : opts.chartColors);

						strXml += '  <c:dPt>';
						strXml += '    <c:idx val="' + index + '"/>';
						strXml += '      <c:invertIfNegative val="' + (opts.invertedColors ? 0 : 1) + '"/>';
						strXml += '    <c:bubble3D val="0"/>';
						strXml += '    <c:spPr>';
						if (opts.lineSize === 0) {
							strXml += '<a:ln><a:noFill/></a:ln>';
						}
						else if (chartType === 'bar') {
							strXml += '<a:solidFill>';
							strXml += '  <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
							strXml += '</a:solidFill>';
						}
						else {
							strXml += '<a:ln>';
							strXml += '  <a:solidFill>';
							strXml += '   <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
							strXml += '  </a:solidFill>';
							strXml += '</a:ln>';
						}
						strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
						strXml += '    </c:spPr>';
						strXml += '  </c:dPt>';
					});
				}

				// 2: "Categories"
				{
					strXml += '<c:cat>';
					if (opts.catLabelFormatCode) {
						// Use 'numRef' as catLabelFormatCode implies that we are expecting numbers here
						strXml += '  <c:numRef>';
						strXml += '    <c:f>Sheet1!' + '$A$2:$A$' + (obj.labels.length + 1) + '</c:f>';
						strXml += '    <c:numCache>';
						strXml += '      <c:formatCode>' + opts.catLabelFormatCode + '</c:formatCode>';
						strXml += '      <c:ptCount val="' + obj.labels.length + '"/>';
						obj.labels.forEach(function(label, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>'; });
						strXml += '    </c:numCache>';
						strXml += '  </c:numRef>';
					}
					else {
						strXml += '  <c:strRef>';
						strXml += '    <c:f>Sheet1!' + '$A$2:$A$' + (obj.labels.length + 1) + '</c:f>';
						strXml += '    <c:strCache>';
						strXml += '	     <c:ptCount val="' + obj.labels.length + '"/>';
						obj.labels.forEach(function(label, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>'; });
						strXml += '    </c:strCache>';
						strXml += '  </c:strRef>';
					}
					strXml += '</c:cat>';
				}

				// 3: "Values"
				{
					strXml += '  <c:val>';
					strXml += '    <c:numRef>';
					strXml += '      <c:f>Sheet1!' + '$' + getExcelColName(idx + 1) + '$2:$' + getExcelColName(idx + 1) + '$' + (obj.labels.length + 1) + '</c:f>';
					strXml += '      <c:numCache>';
					strXml += '        <c:formatCode>General</c:formatCode>';
					strXml += '	       <c:ptCount val="' + obj.labels.length + '"/>';
					obj.values.forEach(function(value, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value == 0 ? value : '') + '</c:v></c:pt>'; });
					strXml += '      </c:numCache>';
					strXml += '    </c:numRef>';
					strXml += '  </c:val>';
				}

				// Option: `smooth`
				if (chartType == 'line') strXml += '<c:smooth val="' + (opts.lineSmooth ? "1" : "0") + '"/>';

				// 4: Close "SERIES"
				strXml += '</c:ser>';
			});

			// 3: "Data Labels"
			{
				strXml += '  <c:dLbls>';
				strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
				strXml += '    <c:txPr>';
				strXml += '      <a:bodyPr/>';
				strXml += '      <a:lstStyle/>';
				strXml += '      <a:p><a:pPr>';
				strXml += '        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
				strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
				strXml += '        </a:defRPr>';
				strXml += '      </a:pPr></a:p>';
				strXml += '    </c:txPr>';
				// NOTE: Throwing an error while creating a multi type chart which contains area chart as the below line appears for the other chart type.
				// Either the given change can be made or the below line can be removed to stop the slide containing multi type chart with area to crash.
				if (opts.type.name != 'area' && opts.type.name != 'radar' && !isMultiTypeChart) strXml += '<c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>';
				strXml += '    <c:showLegendKey val="0"/>';
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
				strXml += '    <c:showCatName val="0"/>';
				strXml += '    <c:showSerName val="0"/>';
				strXml += '    <c:showPercent val="0"/>';
				strXml += '    <c:showBubbleSize val="0"/>';
				strXml += '    <c:showLeaderLines val="0"/>';
				strXml += '  </c:dLbls>';
			}

			// 4: Add more chart options (gapWidth, line Marker, etc.)
			if (chartType == 'bar') {
				strXml += '  <c:gapWidth val="' + opts.barGapWidthPct + '"/>';
				strXml += '  <c:overlap val="' + (opts.barGrouping.indexOf('tacked') > -1 ? 100 : 0) + '"/>';
			}
			else if (chartType == 'bar3D') {
				strXml += '  <c:gapWidth val="' + opts.barGapWidthPct + '"/>';
				strXml += '  <c:gapDepth val="' + opts.barGapDepthPct + '"/>';
				strXml += '  <c:shape val="' + opts.bar3DShape + '"/>';
			}
			else if (chartType == 'line') {
				strXml += '  <c:marker val="1"/>';
			}

			// 5: Add axisId (NOTE: order matters! (category comes first))
			strXml += '  <c:axId val="' + catAxisId + '"/>';
			strXml += '  <c:axId val="' + valAxisId + '"/>';
			strXml += '  <c:axId val="' + AXIS_ID_SERIES_PRIMARY + '"/>';

			// 6: Close Chart tag
			strXml += '</c:' + chartType + 'Chart>';

			// end switch
			break;

		case 'scatter':
			/*
				`data` = [
					{ name:'X-Axis',    values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Value 1', values:[13, 20, 21, 25] },
					{ name:'Y-Value 2', values:[ 1,  2,  5,  9] }
				];
			*/

			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>';
			strXml += '<c:scatterStyle val="lineMarker"/>';
			strXml += '<c:varyColors val="0"/>';

			// 2: Series: (One for each Y-Axis)
			var colorIndex = -1;
			data.filter(function(obj, idx) { return idx > 0; }).forEach(function(obj, idx) {
				colorIndex++;
				strXml += '<c:ser>';
				strXml += '  <c:idx val="' + idx + '"/>';
				strXml += '  <c:order val="' + idx + '"/>';
				strXml += '  <c:tx>';
				strXml += '    <c:strRef>';
				strXml += '      <c:f>Sheet1!$' + LETTERS[(idx + 1)] + '$1</c:f>';
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + obj.name + '</c:v></c:pt></c:strCache>';
				strXml += '    </c:strRef>';
				strXml += '  </c:tx>';

				// 'c:spPr': Fill, Border, Line, LineStyle (dash, etc.), Shadow
				strXml += '  <c:spPr>';
				{
					var strSerColor = opts.chartColors[colorIndex % opts.chartColors.length];

					if (strSerColor == 'transparent') {
						strXml += '<a:noFill/>';
					}
					else if (opts.chartColorsOpacity) {
						strXml += '<a:solidFill>' + createColorElement(strSerColor, '<a:alpha val="' + opts.chartColorsOpacity + '000"/>') + '</a:solidFill>';
					}
					else {
						strXml += '<a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
					}

					if (opts.lineSize == 0) {
						strXml += '<a:ln><a:noFill/></a:ln>';
					}
					else {
						strXml += '<a:ln w="' + (opts.lineSize * ONEPT) + '" cap="flat"><a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
						strXml += '<a:prstDash val="' + (opts.lineDash || "solid") + '"/><a:round/></a:ln>';
					}

					// Shadow
					strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
				}
				strXml += '  </c:spPr>';

				// 'c:marker' tag: `lineDataSymbol`
				{
					strXml += '<c:marker>';
					strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>';
					if (opts.lineDataSymbolSize) {
						// Defaults to "auto" otherwise (but this is usually too small, so there is a default)
						strXml += '  <c:size val="' + opts.lineDataSymbolSize + '"/>';
					}
					strXml += '  <c:spPr>';
					strXml += '    <a:solidFill>' + createColorElement(opts.chartColors[(idx + 1 > opts.chartColors.length ? (Math.floor(Math.random() * opts.chartColors.length)) : idx)]) + '</a:solidFill>';
					var symbolLineColor = opts.lineDataSymbolLineColor || strSerColor;
					strXml += '    <a:ln w="' + opts.lineDataSymbolLineSize + '" cap="flat"><a:solidFill>' + createColorElement(symbolLineColor) + '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
					strXml += '    <a:effectLst/>';
					strXml += '  </c:spPr>';
					strXml += '</c:marker>';
				}

				// Option: scatter data point labels
				if (opts.showLabel) {
					var chartUuid = getUuid('-xxxx-xxxx-xxxx-xxxxxxxxxxxx')
					if (obj.labels && (opts.dataLabelFormatScatter == 'custom' || opts.dataLabelFormatScatter == 'customXY')) {
						strXml += '<c:dLbls>';
						obj.labels.forEach(function(label, idx) {
							if (opts.dataLabelFormatScatter == 'custom' || opts.dataLabelFormatScatter == 'customXY') {
								strXml += '  <c:dLbl>';
								strXml += '    <c:idx val="' + idx + '"/>';
								strXml += '    <c:tx>';
								strXml += '      <c:rich>';
								strXml += '			<a:bodyPr>';
								strXml += '				<a:spAutoFit/>';
								strXml += '			</a:bodyPr>';
								strXml += '        	<a:lstStyle/>';
								strXml += '        	<a:p>';
								strXml += '				<a:pPr>'
								strXml += '					<a:defRPr/>'
								strXml += '				</a:pPr>'
								strXml += '          	<a:r>';
								strXml += '            		<a:rPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>';
								strXml += '            		<a:t>' + encodeXmlEntities(label) + '</a:t>';
								strXml += '          	</a:r>';
								// Apply XY values at end of custom label
								// Do not apply the values if the label was empty or just spaces
								// This allows for selective labelling where required
								if (opts.dataLabelFormatScatter == 'customXY' && !(/^ *$/.test(label))) {
									strXml += '          	<a:r>';
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>';
									strXml += '          		<a:t> (</a:t>';
									strXml += '          	</a:r>';
									strXml += '          	<a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="XVALUE">';
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>';
									strXml += '          		<a:pPr>';
									strXml += '          			<a:defRPr/>';
									strXml += '          		</a:pPr>';
									strXml += '          		<a:t>[' + encodeXmlEntities(obj.name) + '</a:t>';
									strXml += '          	</a:fld>';
									strXml += '          	<a:r>';
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>';
									strXml += '          		<a:t>, </a:t>';
									strXml += '          	</a:r>';
									strXml += '          	<a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="YVALUE">';
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>';
									strXml += '          		<a:pPr>';
									strXml += '          			<a:defRPr/>';
									strXml += '          		</a:pPr>';
									strXml += '          		<a:t>[' + encodeXmlEntities(obj.name) + ']</a:t>';
									strXml += '          	</a:fld>';
									strXml += '          	<a:r>';
									strXml += '          		<a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>';
									strXml += '          		<a:t>)</a:t>';
									strXml += '          	</a:r>';
									strXml += '          	<a:endParaRPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>';
								}
								strXml += '        	</a:p>';
								strXml += '      </c:rich>';
								strXml += '    </c:tx>';
								strXml += '    <c:spPr>';
								strXml += '    	<a:noFill/>';
								strXml += '    	<a:ln>';
								strXml += '    		<a:noFill/>';
								strXml += '    	</a:ln>';
								strXml += '    	<a:effectLst/>';
								strXml += '    </c:spPr>';
								strXml += '    <c:showLegendKey val="0"/>';
								strXml += '    <c:showVal val="0"/>';
								strXml += '    <c:showCatName val="0"/>';
								strXml += '    <c:showSerName val="0"/>';
								strXml += '    <c:showPercent val="0"/>';
								strXml += '    <c:showBubbleSize val="0"/>';
								strXml += '	  <c:showLeaderLines val="1"/>';
								strXml += '    <c:extLst>';
								strXml += '      <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">';
								strXml += '			<c15:dlblFieldTable/>';
								strXml += '			<c15:showDataLabelsRange val="0"/>';
								strXml += '		</c:ext>';
								strXml += '      <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">';
								strXml += '			<c16:uniqueId val="{' + "00000000".substring(0, 8 - (idx + 1).toString().length).toString() + (idx + 1) + chartUuid + '}"/>';
								strXml += '      </c:ext>';
								strXml += '		</c:extLst>';
								strXml += '</c:dLbl>';
							}
						});
						strXml += '</c:dLbls>';
					}
					if (opts.dataLabelFormatScatter == 'XY') {
						strXml += '<c:dLbls>';
						strXml += '	<c:spPr>';
						strXml += '		<a:noFill/>';
						strXml += '		<a:ln>';
						strXml += '			<a:noFill/>';
						strXml += '		</a:ln>';
						strXml += '	  	<a:effectLst/>';
						strXml += '	</c:spPr>';
						strXml += '	<c:txPr>';
						strXml += '		<a:bodyPr>';
						strXml += '			<a:spAutoFit/>';
						strXml += '		</a:bodyPr>';
						strXml += '		<a:lstStyle/>';
						strXml += '		<a:p>';
						strXml += '	    	<a:pPr>';
						strXml += '        		<a:defRPr/>';
						strXml += '	    	</a:pPr>';
						strXml += '	    	<a:endParaRPr lang="en-US"/>';
						strXml += '		</a:p>';
						strXml += '	</c:txPr>';
						strXml += '	<c:showLegendKey val="0"/>';
						strXml += '	<c:showVal val="' + opts.showLabel ? "1" : "0" + '"/>';
						strXml += '	<c:showCatName val="' + opts.showLabel ? "1" : "0" + '"/>';
						strXml += '	<c:showSerName val="0"/>';
						strXml += '	<c:showPercent val="0"/>';
						strXml += '	<c:showBubbleSize val="0"/>';
						strXml += '	<c:extLst>';
						strXml += '		<c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">';
						strXml += '			<c15:showLeaderLines val="1"/>';
						strXml += '		</c:ext>';
						strXml += '	</c:extLst>';
						strXml += '</c:dLbls>';
					}
				}

				// Color bar chart bars various colors
				// Allow users with a single data set to pass their own array of colors (check for this using != ours)
				if ((data.length === 1 || opts.valueBarColors) && opts.chartColors != BARCHART_COLORS) {
					// Series Data Point colors
					obj.values.forEach(function(value, index) {
						var arrColors = (value < 0 ? (opts.invertedColors || BARCHART_COLORS) : opts.chartColors);

						strXml += '  <c:dPt>';
						strXml += '    <c:idx val="' + index + '"/>';
						strXml += '      <c:invertIfNegative val="' + (opts.invertedColors ? 0 : 1) + '"/>';
						strXml += '    <c:bubble3D val="0"/>';
						strXml += '    <c:spPr>';
						if (opts.lineSize === 0) {
							strXml += '<a:ln><a:noFill/></a:ln>';
						}
						else {
							strXml += '<a:solidFill>';
							strXml += ' <a:srgbClr val="' + arrColors[index % arrColors.length] + '"/>';
							strXml += '</a:solidFill>';
						}
						strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
						strXml += '    </c:spPr>';
						strXml += '  </c:dPt>';
					});
				}

				// 3: "Values": Scatter Chart has 2: `xVal` and `yVal`
				{
					// X-Axis is always the same
					strXml += '<c:xVal>';
					strXml += '  <c:numRef>';
					strXml += '    <c:f>Sheet1!$A$2:$A$' + (data[0].values.length + 1) + '</c:f>';
					strXml += '    <c:numCache>';
					strXml += '      <c:formatCode>General</c:formatCode>';
					strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
					data[0].values.forEach(function(value, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value == 0 ? value : '') + '</c:v></c:pt>'; });
					strXml += '    </c:numCache>';
					strXml += '  </c:numRef>';
					strXml += '</c:xVal>';

					// Y-Axis vals are this object's `values`
					strXml += '<c:yVal>';
					strXml += '  <c:numRef>';
					strXml += '    <c:f>Sheet1!$' + getExcelColName(idx + 1) + '$2:$' + getExcelColName(idx + 1) + '$' + (data[0].values.length + 1) + '</c:f>';
					strXml += '    <c:numCache>';
					strXml += '      <c:formatCode>General</c:formatCode>';
					// NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
					strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
					data[0].values.forEach(function(value, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + (obj.values[idx] || obj.values[idx] == 0 ? obj.values[idx] : '') + '</c:v></c:pt>'; });
					strXml += '    </c:numCache>';
					strXml += '  </c:numRef>';
					strXml += '</c:yVal>';
				}

				// Option: `smooth`
				strXml += '<c:smooth val="' + (opts.lineSmooth ? "1" : "0") + '"/>';

				// 4: Close "SERIES"
				strXml += '</c:ser>';
			});

			// 3: Data Labels
			{
				strXml += '  <c:dLbls>';
				strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
				strXml += '    <c:txPr>';
				strXml += '      <a:bodyPr/>';
				strXml += '      <a:lstStyle/>';
				strXml += '      <a:p><a:pPr>';
				strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
				strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
				strXml += '        </a:defRPr>';
				strXml += '      </a:pPr></a:p>';
				strXml += '    </c:txPr>';
				strXml += '    <c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>';
				strXml += '    <c:showLegendKey val="0"/>';
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
				strXml += '    <c:showCatName val="0"/>';
				strXml += '    <c:showSerName val="0"/>';
				strXml += '    <c:showPercent val="0"/>';
				strXml += '    <c:showBubbleSize val="0"/>';
				strXml += '  </c:dLbls>';
			}

			// 4: Add axisId (NOTE: order matters! (category comes first))
			strXml += '  <c:axId val="' + catAxisId + '"/>';
			strXml += '  <c:axId val="' + valAxisId + '"/>';

			// 5: Close Chart tag
			strXml += '</c:' + chartType + 'Chart>';

			// end switch
			break;

		case 'bubble':
			/*
				`data` = [
					{ name:'X-Axis',     values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Values 1', values:[13, 20, 21, 25], sizes:[10, 5, 20, 15] },
					{ name:'Y-Values 2', values:[ 1,  2,  5,  9], sizes:[ 5, 3,  9,  3] }
				];
			*/

			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>';
			strXml += '<c:varyColors val="0"/>';

			// 2: Series: (One for each Y-Axis)
			var colorIndex = -1;
			var idxColLtr = 1;
			data.filter(function(obj, idx) { return idx > 0; }).forEach(function(obj, idx) {
				colorIndex++;
				strXml += '<c:ser>';
				strXml += '  <c:idx val="' + idx + '"/>';
				strXml += '  <c:order val="' + idx + '"/>';

				// A: `<c:tx>`
				strXml += '  <c:tx>';
				strXml += '    <c:strRef>';
				strXml += '      <c:f>Sheet1!$' + LETTERS[idxColLtr] + '$1</c:f>';
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + obj.name + '</c:v></c:pt></c:strCache>';
				strXml += '    </c:strRef>';
				strXml += '  </c:tx>';

				// B: '<c:spPr>': Fill, Border, Line, LineStyle (dash, etc.), Shadow
				{
					strXml += '<c:spPr>';

					var strSerColor = opts.chartColors[colorIndex % opts.chartColors.length];

					if (strSerColor == 'transparent') {
						strXml += '<a:noFill/>';
					}
					else if (opts.chartColorsOpacity) {
						strXml += '<a:solidFill>' + createColorElement(strSerColor, '<a:alpha val="' + opts.chartColorsOpacity + '000"/>') + '</a:solidFill>';
					}
					else {
						strXml += '<a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
					}

					if (opts.lineSize == 0) {
						strXml += '<a:ln><a:noFill/></a:ln>';
					}
					else if (opts.dataBorder) {
						strXml += '<a:ln w="' + (opts.dataBorder.pt * ONEPT) + '" cap="flat"><a:solidFill>' + createColorElement(opts.dataBorder.color) + '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
					}
					else {
						strXml += '<a:ln w="' + (opts.lineSize * ONEPT) + '" cap="flat"><a:solidFill>' + createColorElement(strSerColor) + '</a:solidFill>';
						strXml += '<a:prstDash val="' + (opts.lineDash || "solid") + '"/><a:round/></a:ln>';
					}

					// Shadow
					strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);

					strXml += '</c:spPr>';
				}

				// C: '<c:dLbls>' "Data Labels"
				// Let it be defaulted for now

				// D: '<c:xVal>'/'<c:yVal>' "Values": Scatter Chart has 2: `xVal` and `yVal`
				{
					// X-Axis is always the same
					strXml += '<c:xVal>';
					strXml += '  <c:numRef>';
					strXml += '    <c:f>Sheet1!$A$2:$A$' + (data[0].values.length + 1) + '</c:f>';
					strXml += '    <c:numCache>';
					strXml += '      <c:formatCode>General</c:formatCode>';
					strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
					data[0].values.forEach(function(value, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value == 0 ? value : '') + '</c:v></c:pt>'; });
					strXml += '    </c:numCache>';
					strXml += '  </c:numRef>';
					strXml += '</c:xVal>';

					// Y-Axis vals are this object's `values`
					strXml += '<c:yVal>';
					strXml += '  <c:numRef>';
					strXml += '    <c:f>Sheet1!$' + getExcelColName(idxColLtr) + '$2:$' + getExcelColName(idxColLtr) + '$' + (data[0].values.length + 1) + '</c:f>';
					idxColLtr++;
					strXml += '    <c:numCache>';
					strXml += '      <c:formatCode>General</c:formatCode>';
					// NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
					strXml += '      <c:ptCount val="' + data[0].values.length + '"/>';
					data[0].values.forEach(function(value, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + (obj.values[idx] || obj.values[idx] == 0 ? obj.values[idx] : '') + '</c:v></c:pt>'; });
					strXml += '    </c:numCache>';
					strXml += '  </c:numRef>';
					strXml += '</c:yVal>';
				}

				// E: '<c:bubbleSize>'
				strXml += '  <c:bubbleSize>';
				strXml += '    <c:numRef>';
				strXml += '      <c:f>Sheet1!' + '$' + getExcelColName(idxColLtr) + '$2:$' + getExcelColName(idx + 2) + '$' + (obj.sizes.length + 1) + '</c:f>';
				idxColLtr++;
				strXml += '      <c:numCache>';
				strXml += '        <c:formatCode>General</c:formatCode>';
				strXml += '	       <c:ptCount val="' + obj.sizes.length + '"/>';
				obj.sizes.forEach(function(value, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || '') + '</c:v></c:pt>'; });
				strXml += '      </c:numCache>';
				strXml += '    </c:numRef>';
				strXml += '  </c:bubbleSize>';
				strXml += '  <c:bubble3D val="0"/>';

				// F: Close "SERIES"
				strXml += '</c:ser>';
			});

			// 3: Data Labels
			{
				strXml += '  <c:dLbls>';
				strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
				strXml += '    <c:txPr>';
				strXml += '      <a:bodyPr/>';
				strXml += '      <a:lstStyle/>';
				strXml += '      <a:p><a:pPr>';
				strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
				strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
				strXml += '        </a:defRPr>';
				strXml += '      </a:pPr></a:p>';
				strXml += '    </c:txPr>';
				strXml += '    <c:dLblPos val="ctr"/>';
				strXml += '    <c:showLegendKey val="0"/>';
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
				strXml += '    <c:showCatName val="0"/>';
				strXml += '    <c:showSerName val="0"/>';
				strXml += '    <c:showPercent val="0"/>';
				strXml += '    <c:showBubbleSize val="0"/>';
				strXml += '  </c:dLbls>';
			}

			// 4: Add bubble options
			//strXml += '  <c:bubbleScale val="100"/>';
			//strXml += '  <c:showNegBubbles val="0"/>';
			// Commented out to let it default to PPT until we create options

			// 5: Add axisId (NOTE: order matters! (category comes first))
			strXml += '  <c:axId val="' + catAxisId + '"/>';
			strXml += '  <c:axId val="' + valAxisId + '"/>';

			// 6: Close Chart tag
			strXml += '</c:' + chartType + 'Chart>';

			// end switch
			break;

		case 'pie':
		case 'doughnut':
			// Use the same var name so code blocks from barChart are interchangeable
			var obj = data[0];

			/* EX:
				data: [
				 {
				   name: 'Project Status',
				   labels: ['Red', 'Amber', 'Green', 'Unknown'],
				   values: [10, 20, 38, 2]
				 }
				]
			*/

			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>';
			strXml += '  <c:varyColors val="0"/>';
			strXml += '<c:ser>';
			strXml += '  <c:idx val="0"/>';
			strXml += '  <c:order val="0"/>';
			strXml += '  <c:tx>';
			strXml += '    <c:strRef>';
			strXml += '      <c:f>Sheet1!$B$1</c:f>';
			strXml += '      <c:strCache>';
			strXml += '        <c:ptCount val="1"/>';
			strXml += '        <c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt>';
			strXml += '      </c:strCache>';
			strXml += '    </c:strRef>';
			strXml += '  </c:tx>';
			strXml += '  <c:spPr>';
			strXml += '    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
			strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
			if (opts.dataNoEffects) {
				strXml += '<a:effectLst/>';
			}
			else {
				strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
			}
			strXml += '  </c:spPr>';
			strXml += '<c:explosion val="0"/>';

			// 2: "Data Point" block for every data row
			obj.labels.forEach(function(label, idx) {
				strXml += '<c:dPt>';
				strXml += '  <c:idx val="' + idx + '"/>';
				strXml += '  <c:explosion val="0"/>';
				strXml += '  <c:spPr>';
				strXml += '    <a:solidFill>' + createColorElement(opts.chartColors[(idx + 1 > opts.chartColors.length ? (Math.floor(Math.random() * opts.chartColors.length)) : idx)]) + '</a:solidFill>';
				if (opts.dataBorder) {
					strXml += '<a:ln w="' + (opts.dataBorder.pt * ONEPT) + '" cap="flat"><a:solidFill>' + createColorElement(opts.dataBorder.color) + '</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
				}
				strXml += createShadowElement(opts.shadow, DEF_SHAPE_SHADOW);
				strXml += '  </c:spPr>';
				strXml += '</c:dPt>';
			});

			// 3: "Data Label" block for every data Label
			strXml += '<c:dLbls>';
			obj.labels.forEach(function(label, idx) {
				strXml += '<c:dLbl>';
				strXml += '  <c:idx val="' + idx + '"/>';
				strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
				strXml += '    <c:txPr>';
				strXml += '      <a:bodyPr/><a:lstStyle/>';
				strXml += '      <a:p><a:pPr>';
				strXml += '        <a:defRPr b="' + (opts.dataLabelFontBold ? 1 : 0) + '" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>';
				strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
				strXml += '        </a:defRPr>';
				strXml += '      </a:pPr></a:p>';
				strXml += '    </c:txPr>';
				if (chartType == 'pie') {
					strXml += '    <c:dLblPos val="' + (opts.dataLabelPosition || 'inEnd') + '"/>';
				}
				strXml += '    <c:showLegendKey val="0"/>';
				strXml += '    <c:showVal val="' + (opts.showValue ? "1" : "0") + '"/>';
				strXml += '    <c:showCatName val="' + (opts.showLabel ? "1" : "0") + '"/>';
				strXml += '    <c:showSerName val="0"/>';
				strXml += '    <c:showPercent val="' + (opts.showPercent ? "1" : "0") + '"/>';
				strXml += '    <c:showBubbleSize val="0"/>';
				strXml += '  </c:dLbl>';
			});
			strXml += '<c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>\
				<c:txPr>\
				  <a:bodyPr/>\
				  <a:lstStyle/>\
				  <a:p>\
					<a:pPr>\
					  <a:defRPr b="0" i="0" strike="noStrike" sz="1800" u="none">\
						<a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/>\
					  </a:defRPr>\
					</a:pPr>\
				  </a:p>\
				</c:txPr>\
				' + (chartType == 'pie' ? '<c:dLblPos val="ctr"/>' : '') + '\
				<c:showLegendKey val="0"/>\
				<c:showVal val="0"/>\
				<c:showCatName val="1"/>\
				<c:showSerName val="0"/>\
				<c:showPercent val="1"/>\
				<c:showBubbleSize val="0"/>\
				<c:showLeaderLines val="0"/>';
			strXml += '</c:dLbls>';

			// 2: "Categories"
			strXml += '<c:cat>';
			strXml += '  <c:strRef>';
			strXml += '    <c:f>Sheet1!' + '$A$2:$A$' + (obj.labels.length + 1) + '</c:f>';
			strXml += '    <c:strCache>';
			strXml += '	     <c:ptCount val="' + obj.labels.length + '"/>';
			obj.labels.forEach(function(label, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + encodeXmlEntities(label) + '</c:v></c:pt>'; });
			strXml += '    </c:strCache>';
			strXml += '  </c:strRef>';
			strXml += '</c:cat>';

			// 3: Create vals
			strXml += '  <c:val>';
			strXml += '    <c:numRef>';
			strXml += '      <c:f>Sheet1!' + '$B$2:$B$' + (obj.labels.length + 1) + '</c:f>';
			strXml += '      <c:numCache>';
			strXml += '	       <c:ptCount val="' + obj.labels.length + '"/>';
			obj.values.forEach(function(value, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + (value || value == 0 ? value : '') + '</c:v></c:pt>'; });
			strXml += '      </c:numCache>';
			strXml += '    </c:numRef>';
			strXml += '  </c:val>';

			// 4: Close "SERIES"
			strXml += '  </c:ser>';
			strXml += '  <c:firstSliceAng val="0"/>';
			if (chartType == 'doughnut') strXml += '  <c:holeSize val="' + (opts.holeSize || 50) + '"/>';
			strXml += '</c:' + chartType + 'Chart>';

			// Done with Doughnut/Pie
			break;
	}

	return strXml;
}

function makeCatAxis(opts, axisId, valAxisId) {
	var strXml = '';

	// Build cat axis tag
	// NOTE: Scatter and Bubble chart need two Val axises as they display numbers on x axis
	if (opts.type.name == 'scatter' || opts.type.name == 'bubble') {
		strXml += '<c:valAx>';
	}
	else {
		strXml += '<c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>';
	}
	strXml += '  <c:axId val="' + axisId + '"/>';
	strXml += '  <c:scaling>';
	strXml += '<c:orientation val="' + (opts.catAxisOrientation || (opts.barDir == 'col' ? 'minMax' : 'minMax')) + '"/>';
	if (opts.catAxisMaxVal || opts.catAxisMaxVal == 0) strXml += '<c:max val="' + opts.catAxisMaxVal + '"/>';
	if (opts.catAxisMinVal || opts.catAxisMinVal == 0) strXml += '<c:min val="' + opts.catAxisMinVal + '"/>';
	strXml += '</c:scaling>';
	strXml += '  <c:delete val="' + (opts.catAxisHidden ? 1 : 0) + '"/>';
	strXml += '  <c:axPos val="' + (opts.barDir == 'col' ? 'b' : 'l') + '"/>';
	strXml += (opts.catGridLine !== 'none' ? createGridLineElement(opts.catGridLine, DEF_CHART_GRIDLINE) : '');
	// '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
	if (opts.showCatAxisTitle) {
		strXml += genXmlTitle({
			color: opts.catAxisTitleColor,
			fontFace: opts.catAxisTitleFontFace,
			fontSize: opts.catAxisTitleFontSize,
			rotate: opts.catAxisTitleRotate,
			title: opts.catAxisTitle || 'Axis Title'
		});
	}
	// NOTE: Adding Val Axis Formatting if scatter or bubble charts
	if (opts.type.name == 'scatter' || opts.type.name == 'bubble') {
		strXml += '  <c:numFmt formatCode="' + (opts.valAxisLabelFormatCode ? opts.valAxisLabelFormatCode : 'General') + '" sourceLinked="0"/>';
	}
	else {
		strXml += '  <c:numFmt formatCode="' + (opts.catLabelFormatCode || "General") + '" sourceLinked="0"/>';
	}
	if (opts.type.name === 'scatter') {
		strXml += '  <c:majorTickMark val="none"/>';
		strXml += '  <c:minorTickMark val="none"/>';
		strXml += '  <c:tickLblPos val="nextTo"/>';
	}
	else {
		strXml += '  <c:majorTickMark val="out"/>';
		strXml += '  <c:minorTickMark val="none"/>';
		strXml += '  <c:tickLblPos val="' + (opts.catAxisLabelPos || opts.barDir == 'col' ? 'low' : 'nextTo') + '"/>';
	}
	strXml += '  <c:spPr>';
	strXml += '    <a:ln w="12700" cap="flat">';
	strXml += (opts.catAxisLineShow == false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + DEF_CHART_GRIDLINE.color + '"/></a:solidFill>');
	strXml += '      <a:prstDash val="solid"/>';
	strXml += '      <a:round/>';
	strXml += '    </a:ln>';
	strXml += '  </c:spPr>';
	strXml += '  <c:txPr>';
	strXml += '    <a:bodyPr ' + (opts.catAxisLabelRotate ? ('rot="' + convertRotationDegrees(opts.catAxisLabelRotate) + '"') : "") + '/>'; // don't specify rot 0 so we get the auto behavior
	strXml += '    <a:lstStyle/>';
	strXml += '    <a:p>';
	strXml += '    <a:pPr>';
	strXml += '    <a:defRPr sz="' + (opts.catAxisLabelFontSize || DEF_FONT_SIZE) + '00" b="' + (opts.catAxisLabelFontBold ? 1 : 0) + '" i="0" u="none" strike="noStrike">';
	strXml += '      <a:solidFill><a:srgbClr val="' + (opts.catAxisLabelColor || DEF_FONT_COLOR) + '"/></a:solidFill>';
	strXml += '      <a:latin typeface="' + (opts.catAxisLabelFontFace || 'Arial') + '"/>';
	strXml += '   </a:defRPr>';
	strXml += '  </a:pPr>';
	strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>';
	strXml += '  </a:p>';
	strXml += ' </c:txPr>';
	strXml += ' <c:crossAx val="' + valAxisId + '"/>';
	strXml += ' <c:' + (typeof opts.valAxisCrossesAt === "number" ? 'crossesAt' : 'crosses') + ' val="' + opts.valAxisCrossesAt + '"/>';
	strXml += ' <c:auto val="1"/>';
	strXml += ' <c:lblAlgn val="ctr"/>';
	strXml += ' <c:noMultiLvlLbl val="1"/>';
	if (opts.catAxisLabelFrequency) strXml += ' <c:tickLblSkip val="' + opts.catAxisLabelFrequency + '"/>';

	// Issue#149: PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
	if (opts.catLabelFormatCode) {
		['catAxisBaseTimeUnit', 'catAxisMajorTimeUnit', 'catAxisMinorTimeUnit'].forEach(function(opt, idx) {
			// Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
			if (opts[opt] && (typeof opts[opt] !== 'string' || ['days', 'months', 'years'].indexOf(opt.toLowerCase()) == -1)) {
				console.warn("`" + opt + "` must be one of: 'days','months','years' !");
				opts[opt] = null;
			}
		});
		if (opts.catAxisBaseTimeUnit) strXml += ' <c:baseTimeUnit  val="' + opts.catAxisBaseTimeUnit.toLowerCase() + '"/>';
		if (opts.catAxisMajorTimeUnit) strXml += ' <c:majorTimeUnit val="' + opts.catAxisMajorTimeUnit.toLowerCase() + '"/>';
		if (opts.catAxisMinorTimeUnit) strXml += ' <c:minorTimeUnit val="' + opts.catAxisMinorTimeUnit.toLowerCase() + '"/>';
		if (opts.catAxisMajorUnit) strXml += ' <c:majorUnit     val="' + opts.catAxisMajorUnit + '"/>';
		if (opts.catAxisMinorUnit) strXml += ' <c:minorUnit     val="' + opts.catAxisMinorUnit + '"/>';
	}

	// Close cat axis tag
	// NOTE: Added closing tag of val or cat axis based on chart type
	if (opts.type.name == 'scatter' || opts.type.name == 'bubble') {
		strXml += '</c:valAx>';
	}
	else {
		strXml += '</c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>';
	}

	return strXml;
}

function makeValueAxis(opts, valAxisId) {
	var axisPos = valAxisId === AXIS_ID_VALUE_PRIMARY ? (opts.barDir == 'col' ? 'l' : 'b') : (opts.barDir == 'col' ? 'r' : 't');
	var strXml = '';
	var isRight = axisPos === 'r' || axisPos === 't';
	var crosses = isRight ? 'max' : 'autoZero';
	var crossAxId = valAxisId === AXIS_ID_VALUE_PRIMARY ? AXIS_ID_CATEGORY_PRIMARY : AXIS_ID_CATEGORY_SECONDARY;

	strXml += '<c:valAx>';
	strXml += '  <c:axId val="' + valAxisId + '"/>';
	strXml += '  <c:scaling>';
	strXml += '    <c:orientation val="' + (opts.valAxisOrientation || (opts.barDir == 'col' ? 'minMax' : 'minMax')) + '"/>';
	if (opts.valAxisMaxVal || opts.valAxisMaxVal == 0) strXml += '<c:max val="' + opts.valAxisMaxVal + '"/>';
	if (opts.valAxisMinVal || opts.valAxisMinVal == 0) strXml += '<c:min val="' + opts.valAxisMinVal + '"/>';
	strXml += '  </c:scaling>';
	strXml += '  <c:delete val="' + (opts.valAxisHidden ? 1 : 0) + '"/>';
	strXml += '  <c:axPos val="' + axisPos + '"/>';
	if (opts.valGridLine != 'none') strXml += createGridLineElement(opts.valGridLine, DEF_CHART_GRIDLINE);
	// '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
	if (opts.showValAxisTitle) {
		strXml += genXmlTitle({
			color: opts.valAxisTitleColor,
			fontFace: opts.valAxisTitleFontFace,
			fontSize: opts.valAxisTitleFontSize,
			rotate: opts.valAxisTitleRotate,
			title: opts.valAxisTitle || 'Axis Title'
		});
	}
	strXml += ' <c:numFmt formatCode="' + (opts.valAxisLabelFormatCode ? opts.valAxisLabelFormatCode : 'General') + '" sourceLinked="0"/>';
	if (opts.type.name === 'scatter') {
		strXml += '  <c:majorTickMark val="none"/>';
		strXml += '  <c:minorTickMark val="none"/>';
		strXml += '  <c:tickLblPos val="nextTo"/>';
	}
	else {
		strXml += ' <c:majorTickMark val="out"/>';
		strXml += ' <c:minorTickMark val="none"/>';
		strXml += ' <c:tickLblPos val="' + (opts.catAxisLabelPos || opts.barDir == 'col' ? 'nextTo' : 'low') + '"/>';
	}
	strXml += ' <c:spPr>';
	strXml += '   <a:ln w="12700" cap="flat">';
	strXml += (opts.valAxisLineShow == false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + DEF_CHART_GRIDLINE.color + '"/></a:solidFill>');
	strXml += '     <a:prstDash val="solid"/>';
	strXml += '     <a:round/>';
	strXml += '   </a:ln>';
	strXml += ' </c:spPr>';
	strXml += ' <c:txPr>';
	strXml += '  <a:bodyPr ' + (opts.valAxisLabelRotate ? ('rot="' + convertRotationDegrees(opts.valAxisLabelRotate) + '"') : "") + '/>'; // don't specify rot 0 so we get the auto behavior
	strXml += '  <a:lstStyle/>';
	strXml += '  <a:p>';
	strXml += '    <a:pPr>';
	strXml += '      <a:defRPr sz="' + (opts.valAxisLabelFontSize || DEF_FONT_SIZE) + '00" b="' + (opts.valAxisLabelFontBold ? 1 : 0) + '" i="0" u="none" strike="noStrike">';
	strXml += '        <a:solidFill><a:srgbClr val="' + (opts.valAxisLabelColor || DEF_FONT_COLOR) + '"/></a:solidFill>';
	strXml += '        <a:latin typeface="' + (opts.valAxisLabelFontFace || 'Arial') + '"/>';
	strXml += '      </a:defRPr>';
	strXml += '    </a:pPr>';
	strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>';
	strXml += '  </a:p>';
	strXml += ' </c:txPr>';
	strXml += ' <c:crossAx val="' + crossAxId + '"/>';
	strXml += ' <c:crosses val="' + crosses + '"/>';
	strXml += ' <c:crossBetween val="' + (opts.type.name === 'scatter' || opts.hasArea ? 'midCat' : 'between') + '"/>';
	if (opts.valAxisMajorUnit) strXml += ' <c:majorUnit val="' + opts.valAxisMajorUnit + '"/>';
	strXml += '</c:valAx>';

	return strXml;
}

/** DESC: Used by `bar3D` */
function makeSerAxis(opts, axisId, valAxisId) {
	var strXml = '';

	// Build ser axis tag
	strXml += '<c:serAx>';
	strXml += '  <c:axId val="' + axisId + '"/>';
	strXml += '  <c:scaling><c:orientation val="' + (opts.serAxisOrientation || (opts.barDir == 'col' ? 'minMax' : 'minMax')) + '"/></c:scaling>';
	strXml += '  <c:delete val="' + (opts.serAxisHidden ? 1 : 0) + '"/>';
	strXml += '  <c:axPos val="' + (opts.barDir == 'col' ? 'b' : 'l') + '"/>';
	strXml += (opts.serGridLine !== 'none' ? createGridLineElement(opts.serGridLine, DEF_CHART_GRIDLINE) : '');
	// '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
	if (opts.showSerAxisTitle) {
		strXml += genXmlTitle({
			color: opts.serAxisTitleColor,
			fontFace: opts.serAxisTitleFontFace,
			fontSize: opts.serAxisTitleFontSize,
			rotate: opts.serAxisTitleRotate,
			title: opts.serAxisTitle || 'Axis Title'
		});
	}
	strXml += '  <c:numFmt formatCode="' + (opts.serLabelFormatCode || "General") + '" sourceLinked="0"/>';
	strXml += '  <c:majorTickMark val="out"/>';
	strXml += '  <c:minorTickMark val="none"/>';
	strXml += '  <c:tickLblPos val="' + (opts.serAxisLabelPos || opts.barDir == 'col' ? 'low' : 'nextTo') + '"/>';
	strXml += '  <c:spPr>';
	strXml += '    <a:ln w="12700" cap="flat">';
	strXml += (opts.serAxisLineShow == false ? '<a:noFill/>' : '<a:solidFill><a:srgbClr val="' + DEF_CHART_GRIDLINE.color + '"/></a:solidFill>');
	strXml += '      <a:prstDash val="solid"/>';
	strXml += '      <a:round/>';
	strXml += '    </a:ln>';
	strXml += '  </c:spPr>';
	strXml += '  <c:txPr>';
	strXml += '    <a:bodyPr/>';  // don't specify rot 0 so we get the auto behavior
	strXml += '    <a:lstStyle/>';
	strXml += '    <a:p>';
	strXml += '    <a:pPr>';
	strXml += '    <a:defRPr sz="' + (opts.serAxisLabelFontSize || DEF_FONT_SIZE) + '00" b="0" i="0" u="none" strike="noStrike">';
	strXml += '      <a:solidFill><a:srgbClr val="' + (opts.serAxisLabelColor || DEF_FONT_COLOR) + '"/></a:solidFill>';
	strXml += '      <a:latin typeface="' + (opts.serAxisLabelFontFace || 'Arial') + '"/>';
	strXml += '   </a:defRPr>';
	strXml += '  </a:pPr>';
	strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>';
	strXml += '  </a:p>';
	strXml += ' </c:txPr>';
	strXml += ' <c:crossAx val="' + valAxisId + '"/>';
	strXml += ' <c:crosses val="autoZero"/>';
	if (opts.serAxisLabelFrequency) strXml += ' <c:tickLblSkip val="' + opts.serAxisLabelFrequency + '"/>';

	// Issue#149: PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
	if (opts.serLabelFormatCode) {
		['serAxisBaseTimeUnit', 'serAxisMajorTimeUnit', 'serAxisMinorTimeUnit'].forEach(function(opt, idx) {
			// Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
			if (opts[opt] && (typeof opts[opt] !== 'string' || ['days', 'months', 'years'].indexOf(opt.toLowerCase()) == -1)) {
				console.warn("`" + opt + "` must be one of: 'days','months','years' !");
				opts[opt] = null;
			}
		});
		if (opts.serAxisBaseTimeUnit) strXml += ' <c:baseTimeUnit  val="' + opts.serAxisBaseTimeUnit.toLowerCase() + '"/>';
		if (opts.serAxisMajorTimeUnit) strXml += ' <c:majorTimeUnit val="' + opts.serAxisMajorTimeUnit.toLowerCase() + '"/>';
		if (opts.serAxisMinorTimeUnit) strXml += ' <c:minorTimeUnit val="' + opts.serAxisMinorTimeUnit.toLowerCase() + '"/>';
		if (opts.serAxisMajorUnit) strXml += ' <c:majorUnit     val="' + opts.serAxisMajorUnit + '"/>';
		if (opts.serAxisMinorUnit) strXml += ' <c:minorUnit     val="' + opts.serAxisMinorUnit + '"/>';
	}

	// Close ser axis tag
	strXml += '</c:serAx>';

	return strXml;
}

/**
* DESC: Convert degrees (0..360) to Powerpoint rot value
*/
function convertRotationDegrees(d) {
	d = d || 0;
	return (d > 360 ? (d - 360) : d) * 60000;
}

/**
* DESC: Generate the XML for title elements used for the char and axis titles
*/
function genXmlTitle(opts) {
	var align = (opts.titleAlign == 'left' ? 'l' : (opts.titleAlign == 'right' ? 'r' : false));
	var strXml = '';
	strXml += '<c:title>';
	strXml += ' <c:tx>';
	strXml += '  <c:rich>';
	if (opts.rotate) {
		strXml += '  <a:bodyPr rot="' + convertRotationDegrees(opts.rotate) + '"/>';
	}
	else {
		strXml += '  <a:bodyPr/>';  // don't specify rotation to get default (ex. vertical for cat axis)
	}
	strXml += '  <a:lstStyle/>';
	strXml += '  <a:p>';
	strXml += (align ? '<a:pPr algn="' + align + '">' : '<a:pPr>');
	var sizeAttr = '';
	if (opts.fontSize) {
		// only set the font size if specified.  Powerpoint will handle the default size
		sizeAttr = 'sz="' + Math.round(opts.fontSize) + '00"';
	}
	strXml += '      <a:defRPr ' + sizeAttr + ' b="0" i="0" u="none" strike="noStrike">';
	strXml += '        <a:solidFill><a:srgbClr val="' + (opts.color || DEF_FONT_COLOR) + '"/></a:solidFill>';
	strXml += '        <a:latin typeface="' + (opts.fontFace || 'Arial') + '"/>';
	strXml += '      </a:defRPr>';
	strXml += '    </a:pPr>';
	strXml += '    <a:r>';
	strXml += '      <a:rPr ' + sizeAttr + ' b="0" i="0" u="none" strike="noStrike">';
	strXml += '        <a:solidFill><a:srgbClr val="' + (opts.color || DEF_FONT_COLOR) + '"/></a:solidFill>';
	strXml += '        <a:latin typeface="' + (opts.fontFace || 'Arial') + '"/>';
	strXml += '      </a:rPr>';
	strXml += '      <a:t>' + (encodeXmlEntities(opts.title) || '') + '</a:t>';
	strXml += '    </a:r>';
	strXml += '  </a:p>';
	strXml += '  </c:rich>';
	strXml += ' </c:tx>';
	if (opts.titlePos && opts.titlePos.x && opts.titlePos.y) {
		strXml += '<c:layout>';
		strXml += '  <c:manualLayout>';
		strXml += '    <c:xMode val="edge"/>';
		strXml += '    <c:yMode val="edge"/>';
		strXml += '    <c:x val="' + opts.titlePos.x + '"/>';
		strXml += '    <c:y val="' + opts.titlePos.y + '"/>';
		strXml += '  </c:manualLayout>';
		strXml += '</c:layout>';
	}
	else {
		strXml += ' <c:layout/>';
	}
	strXml += ' <c:overlay val="0"/>';
	strXml += '</c:title>';
	return strXml;
}

/**
* DESC: Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
* EX:
	<p:txBody>
		<a:bodyPr wrap="none" lIns="50800" tIns="50800" rIns="50800" bIns="50800" anchor="ctr">
		</a:bodyPr>
		<a:lstStyle/>
		<a:p>
		  <a:pPr marL="228600" indent="-228600"><a:buSzPct val="100000"/><a:buChar char="&#x2022;"/></a:pPr>
		  <a:r>
			<a:t>bullet 1 </a:t>
		  </a:r>
		  <a:r>
			<a:rPr>
			  <a:solidFill><a:srgbClr val="7B2CD6"/></a:solidFill>
			</a:rPr>
			<a:t>colored text</a:t>
		  </a:r>
		</a:p>
	  </p:txBody>
* NOTES:
* - PPT text lines [lines followed by line-breaks] are createing using <p>-aragraph's
* - Bullets are a paragprah-level formatting device
*
* @param slideObj (object) - slideObj -OR- table `cell` object
* @returns XML string containing the param object's text and formatting
*/
function genXmlTextBody(slideObj) {
	// FIRST: Shapes without text, etc. may be sent here during build, but have no text to render so return an empty string
	if (!slideObj.options.isTableCell && (typeof slideObj.text === 'undefined' || slideObj.text == null)) return '';

	// Create options if needed
	if (!slideObj.options) slideObj.options = {};

	// Vars
	var arrTextObjects = [];
	var tagStart = (slideObj.options.isTableCell ? '<a:txBody>' : '<p:txBody>');
	var tagClose = (slideObj.options.isTableCell ? '</a:txBody>' : '</p:txBody>');
	var strSlideXml = tagStart;

	// STEP 1: Modify slideObj to be consistent array of `{ text:'', options:{} }`
	/* CASES:
		addText( 'string' )
		addText( 'line1\n line2' )
		addText( ['barry','allen'] )
		addText( [{text'word1'}, {text:'word2'}] )
		addText( [{text'line1\n line2'}, {text:'end word'}] )
	*/
	// A: Handle string/number
	if (typeof slideObj.text === 'string' || typeof slideObj.text === 'number') {
		slideObj.text = [{ text: slideObj.text.toString(), options: (slideObj.options || {}) }];
	}

	// STEP 2: Grab options, format line-breaks, etc.
	if (Array.isArray(slideObj.text)) {
		slideObj.text.forEach(function(obj, idx) {
			// A: Set options
			obj.options = obj.options || slideObj.options || {};
			if (idx == 0 && obj.options && !obj.options.bullet && slideObj.options.bullet) obj.options.bullet = slideObj.options.bullet;

			// B: Cast to text-object and fix line-breaks (if needed)
			if (typeof obj.text === 'string' || typeof obj.text === 'number') {
				obj.text = obj.text.toString().replace(/\r*\n/g, CRLF);
				// Plain strings like "hello \n world" need to have lineBreaks set to break as intended
				if (obj.text.indexOf(CRLF) > -1) obj.options.breakLine = true;
			}

			// C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
			if (obj.text.split(CRLF).length > 0) {
				obj.text.toString().split(CRLF).forEach(function(line, idx) {
					// Add line-breaks if not bullets/aligned (we add CRLF for those below in STEP 2)
					line += (obj.options.breakLine && !obj.options.bullet && !obj.options.align ? CRLF : '');
					arrTextObjects.push({ text: line, options: obj.options });
				});
			}
			else {
				// NOTE: The replace used here is for non-textObjects (plain strings) eg:'hello\nworld'
				arrTextObjects.push(obj);
			}
		});
	}

	// STEP 3: Add bodyProperties
	{
		// A: 'bodyPr'
		strSlideXml += genXmlBodyProperties(slideObj.options);

		// B: 'lstStyle'
		// NOTE: Shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
		// FIXME: LINE horiz-align doesnt work (text is always to the left inside line) (FYI: the PPT code diff is substantial!)
		if (slideObj.options.h == 0 && slideObj.options.line && slideObj.options.align) {
			strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>';
		}
		else if (slideObj.type === 'placeholder') {
			strSlideXml += '<a:lstStyle>';
			strSlideXml += genXmlParagraphProperties(slideObj, true);
			strSlideXml += '</a:lstStyle>';
		}
		else {
			strSlideXml += '<a:lstStyle/>';
		}
	}

	// STEP 4: Loop over each text object and create paragraph props, text run, etc.
	arrTextObjects.forEach(function(textObj, idx) {
		// Clear/Increment loop vars
		paragraphPropXml = '<a:pPr ' + (textObj.options.rtlMode ? ' rtl="1" ' : '');
		strXmlBullet = '', strXmlParaSpc = '';
		textObj.options.lineIdx = idx;

		// Inherit pPr-type options from parent shape's `options`
		textObj.options.align = textObj.options.align || slideObj.options.align;
		textObj.options.lineSpacing = textObj.options.lineSpacing || slideObj.options.lineSpacing;
		textObj.options.indentLevel = textObj.options.indentLevel || slideObj.options.indentLevel;
		textObj.options.paraSpaceBefore = textObj.options.paraSpaceBefore || slideObj.options.paraSpaceBefore;
		textObj.options.paraSpaceAfter = textObj.options.paraSpaceAfter || slideObj.options.paraSpaceAfter;

		textObj.options.lineIdx = idx;
		var paragraphPropXml = genXmlParagraphProperties(textObj, false);

		// B: Start paragraph if this is the first text obj, or if current textObj is about to be bulleted or aligned
		if (idx == 0) {
			// Add paragraphProperties right after <p> before textrun(s) begin
			strSlideXml += '<a:p>' + paragraphPropXml;
		}
		else if (idx > 0 && (typeof textObj.options.bullet !== 'undefined' || typeof textObj.options.align !== 'undefined')) {
			strSlideXml += '</a:p><a:p>' + paragraphPropXml;
		}

		// C: Inherit any main options (color, fontSize, etc.)
		// We only pass the text.options to genXmlTextRun (not the Slide.options),
		// so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
		jQuery.each(slideObj.options, function(key, val) {
			// NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
			if (key != 'bullet' && !textObj.options[key]) textObj.options[key] = val;
		});

		// D: Add formatted textrun
		strSlideXml += genXmlTextRun(textObj.options, textObj.text);
	});

	// STEP 5: Append 'endParaRPr' (when needed) and close current open paragraph
	// NOTE: (ISSUE#20/#193): Add 'endParaRPr' with font/size props or PPT default (Arial/18pt en-us) is used making row "too tall"/not honoring opts
	if (slideObj.options.isTableCell && (slideObj.options.fontSize || slideObj.options.fontFace)) {
		strSlideXml += '<a:endParaRPr lang="' + (slideObj.options.lang ? slideObj.options.lang : 'en-US') + '" '
			+ (slideObj.options.fontSize ? ' sz="' + Math.round(slideObj.options.fontSize) + '00"' : '') + ' dirty="0">';
		if (slideObj.options.fontFace) {
			strSlideXml += '  <a:latin typeface="' + slideObj.options.fontFace + '" charset="0" />';
			strSlideXml += '  <a:ea    typeface="' + slideObj.options.fontFace + '" charset="0" />';
			strSlideXml += '  <a:cs    typeface="' + slideObj.options.fontFace + '" charset="0" />';
		}
		strSlideXml += '</a:endParaRPr>';
	}
	else {
		strSlideXml += '<a:endParaRPr lang="' + (slideObj.options.lang || 'en-US') + '" dirty="0"/>'; // NOTE: Added 20180101 to address PPT-2007 issues
	}
	strSlideXml += '</a:p>';

	// STEP 6: Close the textBody
	strSlideXml += tagClose;

	// LAST: Return XML
	return strSlideXml;
}

function genXmlParagraphProperties(textObj, isDefault) {
	var strXmlBullet = '', strXmlLnSpc = '', strXmlParaSpc = '';
	var bulletLvl0Margin = 342900;
	var tag = isDefault ? 'a:lvl1pPr' : 'a:pPr';

	var paragraphPropXml = '<' + tag + (textObj.options.rtlMode ? ' rtl="1" ' : '');

	// A: Build paragraphProperties
	{
		// OPTION: align
		if (textObj.options.align) {
			switch (textObj.options.align) {
				case 'l':
				case 'left':
					paragraphPropXml += ' algn="l"';
					break;
				case 'r':
				case 'right':
					paragraphPropXml += ' algn="r"';
					break;
				case 'c':
				case 'ctr':
				case 'center':
					paragraphPropXml += ' algn="ctr"';
					break;
				case 'justify':
					paragraphPropXml += ' algn="just"';
					break;
			}
		}

		if (textObj.options.lineSpacing) {
			strXmlLnSpc = '<a:lnSpc><a:spcPts val="' + textObj.options.lineSpacing + '00"/></a:lnSpc>';
		}

		// OPTION: indent
		if (textObj.options.indentLevel && !isNaN(Number(textObj.options.indentLevel)) && textObj.options.indentLevel > 0) {
			paragraphPropXml += ' lvl="' + textObj.options.indentLevel + '"';
		}

		// OPTION: Paragraph Spacing: Before/After
		if (textObj.options.paraSpaceBefore && !isNaN(Number(textObj.options.paraSpaceBefore)) && textObj.options.paraSpaceBefore > 0) {
			strXmlParaSpc += '<a:spcBef><a:spcPts val="' + (textObj.options.paraSpaceBefore * 100) + '"/></a:spcBef>';
		}
		if (textObj.options.paraSpaceAfter && !isNaN(Number(textObj.options.paraSpaceAfter)) && textObj.options.paraSpaceAfter > 0) {
			strXmlParaSpc += '<a:spcAft><a:spcPts val="' + (textObj.options.paraSpaceAfter * 100) + '"/></a:spcAft>';
		}

		// Set core XML for use below
		paraPropXmlCore = paragraphPropXml;

		// OPTION: bullet
		// NOTE: OOXML uses the unicode character set for Bullets
		// EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
		if (typeof textObj.options.bullet === 'object') {
			if (textObj.options.bullet.type) {
				if (textObj.options.bullet.type.toString().toLowerCase() == "number") {
					paragraphPropXml += ' marL="' + (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin + (bulletLvl0Margin * textObj.options.indentLevel) : bulletLvl0Margin) + '" indent="-' + bulletLvl0Margin + '"';
					strXmlBullet = '<a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>';
				}
			}
			else if (textObj.options.bullet.code) {
				var bulletCode = '&#x' + textObj.options.bullet.code + ';';

				// Check value for hex-ness (s/b 4 char hex)
				if (/^[0-9A-Fa-f]{4}$/.test(textObj.options.bullet.code) == false) {
					console.warn('Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!');
					bulletCode = BULLET_TYPES['DEFAULT'];
				}

				paragraphPropXml += ' marL="' + (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin + (bulletLvl0Margin * textObj.options.indentLevel) : bulletLvl0Margin) + '" indent="-' + bulletLvl0Margin + '"';
				strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + bulletCode + '"/>';
			}
		}
		else if (textObj.options.bullet == true) {
			paragraphPropXml += ' marL="' + (textObj.options.indentLevel && textObj.options.indentLevel > 0 ? bulletLvl0Margin + (bulletLvl0Margin * textObj.options.indentLevel) : bulletLvl0Margin) + '" indent="-' + bulletLvl0Margin + '"';
			strXmlBullet = '<a:buSzPct val="100000"/><a:buChar char="' + BULLET_TYPES['DEFAULT'] + '"/>';
		}
		else {
			strXmlBullet = '<a:buNone/>';
		}

		// Close Paragraph-Properties --------------------
		// IMPORTANT: strXmlLnSpc, strXmlParaSpc, and strXmlBullet require strict ordering.
		//            anything out of order is ignored. (PPT-Online, PPT for Mac)
		paragraphPropXml += '>' + strXmlLnSpc + strXmlParaSpc + strXmlBullet;
		if (isDefault) {
			paragraphPropXml += genXmlTextRunProperties(textObj.options, true);
		}
		paragraphPropXml += '</' + tag + '>';
	}

	return paragraphPropXml;
}

function genXmlTextRunProperties(opts, isDefault) {
	var runProps = '';
	var runPropsTag = isDefault ? 'a:defRPr' : 'a:rPr';

	// BEGIN runProperties
	runProps += '<' + runPropsTag + ' lang="' + (opts.lang ? opts.lang : 'en-US') + '" ' + (opts.lang ? ' altLang="en-US"' : '');
	runProps += (opts.bold ? ' b="1"' : '');
	runProps += (opts.fontSize ? ' sz="' + Math.round(opts.fontSize) + '00"' : ''); // NOTE: Use round so sizes like '7.5' wont cause corrupt pres.
	runProps += (opts.italic ? ' i="1"' : '');
	runProps += (opts.strike ? ' strike="sngStrike"' : '');
	runProps += (opts.underline || opts.hyperlink ? ' u="sng"' : '');
	runProps += (opts.subscript ? ' baseline="-40000"' : (opts.superscript ? ' baseline="30000"' : ''));
	runProps += (opts.charSpacing ? ' spc="' + (opts.charSpacing * 100) + '" kern="0"' : ''); // IMPORTANT: Also disable kerning; otherwise text won't actually expand
	runProps += ' dirty="0" smtClean="0">';
	// Color / Font / Outline are children of <a:rPr>, so add them now before closing the runProperties tag
	if (opts.color || opts.fontFace || opts.outline) {
		if (opts.outline && typeof opts.outline === 'object') {
			runProps += ('<a:ln w="' + Math.round((opts.outline.size || 0.75) * ONEPT) + '">' + genXmlColorSelection(opts.outline.color || 'FFFFFF') + '</a:ln>');
		}
		if (opts.color) runProps += genXmlColorSelection(opts.color);
		if (opts.fontFace) {
			// NOTE: 'cs' = Complex Script, 'ea' = East Asian (use -120 instead of 0 - see Issue #174); ea must come first (see Issue #174)
			runProps += '<a:latin typeface="' + opts.fontFace + '" pitchFamily="34" charset="0" />'
				+ '<a:ea typeface="' + opts.fontFace + '" pitchFamily="34" charset="-122" />'
				+ '<a:cs typeface="' + opts.fontFace + '" pitchFamily="34" charset="-120" />';
		}
	}

	// Hyperlink support
	if (opts.hyperlink) {
		if (typeof opts.hyperlink !== 'object') console.log("ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:'https://github.com'}` ");
		else if (!opts.hyperlink.url && !opts.hyperlink.slide) console.log("ERROR: 'hyperlink requires either `url` or `slide`'");
		else if (opts.hyperlink.url) {
			// FIXME-20170410: FUTURE-FEATURE: color (link is always blue in Keynote and PPT online, so usual text run above isnt honored for links..?)
			//runProps += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010! (Issue#74)
			runProps += '<a:hlinkClick r:id="rId' + opts.hyperlink.rId + '" invalidUrl="" action="" tgtFrame="" tooltip="' + (opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') + '" history="1" highlightClick="0" endSnd="0" />';
		}
		else if (opts.hyperlink.slide) {
			runProps += '<a:hlinkClick r:id="rId' + opts.hyperlink.rId + '" action="ppaction://hlinksldjump" tooltip="' + (opts.hyperlink.tooltip ? encodeXmlEntities(opts.hyperlink.tooltip) : '') + '" />';
		}
	}

	// END runProperties
	runProps += '</' + runPropsTag + '>';

	return runProps;
}

/**
* DESC: Builds <a:r></a:r> text runs for <a:p> paragraphs in textBody
* EX:
<a:r>
  <a:rPr lang="en-US" sz="2800" dirty="0" smtClean="0">
	<a:solidFill>
	  <a:srgbClr val="00FF00">
	  </a:srgbClr>
	</a:solidFill>
	<a:latin typeface="Courier New" pitchFamily="34" charset="0"/>
  </a:rPr>
  <a:t>Misc font/color, size = 28</a:t>
</a:r>
*/
function genXmlTextRun(opts, inStrText) {
	var xmlTextRun = '';
	var paraProp = '';
	var parsedText;

	// ADD runProperties
	var startInfo = genXmlTextRunProperties(opts, false);

	// LINE-BREAKS/MULTI-LINE: Split text into multi-p:
	parsedText = inStrText.split(CRLF);
	if (parsedText.length > 1) {
		var outTextData = '';
		for (var i = 0, total_size_i = parsedText.length; i < total_size_i; i++) {
			outTextData += '<a:r>' + startInfo + '<a:t>' + encodeXmlEntities(parsedText[i]);
			// Stop/Start <p>aragraph as long as there is more lines ahead (otherwise its closed at the end of this function)
			if ((i + 1) < total_size_i) outTextData += (opts.breakLine ? CRLF : '') + '</a:t></a:r>';
		}
		xmlTextRun = outTextData;
	}
	else {
		// Handle cases where addText `text` was an array of objects - if a text object doesnt contain a '\n' it still need alignment!
		// The first pPr-align is done in makeXml - use line countr to ensure we only add subsequently as needed
		xmlTextRun = ((opts.align && opts.lineIdx > 0) ? paraProp : '') + '<a:r>' + startInfo + '<a:t>' + encodeXmlEntities(inStrText);
	}

	// Return paragraph with text run
	return xmlTextRun + '</a:t></a:r>';
}

/**
* DESC: Builds <a:bodyPr></a:bodyPr> tag
*/
function genXmlBodyProperties(objOptions) {
	var bodyProperties = '<a:bodyPr';

	if (objOptions && objOptions.bodyProp) {
		// A: Enable or disable textwrapping none or square:
		(objOptions.bodyProp.wrap) ? bodyProperties += ' wrap="' + objOptions.bodyProp.wrap + '" rtlCol="0"' : bodyProperties += ' wrap="square" rtlCol="0"';

		// B: Set anchorPoints:
		if (objOptions.bodyProp.anchor) bodyProperties += ' anchor="' + objOptions.bodyProp.anchor + '"'; // VALS: [t,ctr,b]
		if (objOptions.bodyProp.vert) bodyProperties += ' vert="' + objOptions.bodyProp.vert + '"'; // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]

		// C: Textbox margins [padding]:
		if (objOptions.bodyProp.bIns || objOptions.bodyProp.bIns == 0) bodyProperties += ' bIns="' + objOptions.bodyProp.bIns + '"';
		if (objOptions.bodyProp.lIns || objOptions.bodyProp.lIns == 0) bodyProperties += ' lIns="' + objOptions.bodyProp.lIns + '"';
		if (objOptions.bodyProp.rIns || objOptions.bodyProp.rIns == 0) bodyProperties += ' rIns="' + objOptions.bodyProp.rIns + '"';
		if (objOptions.bodyProp.tIns || objOptions.bodyProp.tIns == 0) bodyProperties += ' tIns="' + objOptions.bodyProp.tIns + '"';

		// D: Close <a:bodyPr element
		bodyProperties += '>';

		// E: NEW: Add autofit type tags
		if (objOptions.shrinkText) bodyProperties += '<a:normAutofit fontScale="85000" lnSpcReduction="20000" />'; // MS-PPT > Format Shape > Text Options: "Shrink text on overflow"
		// MS-PPT > Format Shape > Text Options: "Resize shape to fit text" [spAutoFit]
		// NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
		bodyProperties += (objOptions.bodyProp.autoFit !== false ? '<a:spAutoFit/>' : '');

		// LAST: Close bodyProp
		bodyProperties += '</a:bodyPr>';
	}
	else {
		// DEFAULT:
		bodyProperties += ' wrap="square" rtlCol="0">';
		bodyProperties += '</a:bodyPr>';
	}

	// LAST: Return Close bodyProp
	return (objOptions.isTableCell ? '<a:bodyPr/>' : bodyProperties);
}

function genXmlColorSelection(color_info, back_info?: string) {
	var colorVal;
	var fillType = 'solid';
	var internalElements = '';
	var outText = '';

	if (back_info && typeof back_info === 'string') {
		outText += '<p:bg><p:bgPr>';
		outText += genXmlColorSelection(back_info.replace('#', ''));
		outText += '<a:effectLst/>';
		outText += '</p:bgPr></p:bg>';
	}

	if (color_info) {
		if (typeof color_info == 'string') colorVal = color_info;
		else {
			if (color_info.type) fillType = color_info.type;
			if (color_info.color) colorVal = color_info.color;
			if (color_info.alpha) internalElements += '<a:alpha val="' + (100 - color_info.alpha) + '000"/>';
		}

		switch (fillType) {
			case 'solid':
				outText += '<a:solidFill>' + createColorElement(colorVal, internalElements) + '</a:solidFill>';
				break;
		}
	}

	return outText;
}

function genXmlPlaceholder(placeholderObj) {
	var strXml = '';

	if (placeholderObj) {
		var placeholderIdx = placeholderObj.options && placeholderObj.options.placeholderIdx ? placeholderObj.options.placeholderIdx : '';
		var placeholderType = placeholderObj.options && placeholderObj.options.placeholderType ? placeholderObj.options.placeholderType : '';

		strXml += '<p:ph'
			+ (placeholderIdx ? ' idx="' + placeholderIdx + '"' : '')
			+ (placeholderType && PLACEHOLDER_TYPES[placeholderType] ? ' type="' + PLACEHOLDER_TYPES[placeholderType] + '"' : '')
			+ (placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : '')
			+ '/>';
	}
	return strXml;
}

// XML-GEN: First 6 functions create the base /ppt files

function makeXmlContTypes() {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
	strXml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
	strXml += ' <Default Extension="xml" ContentType="application/xml"/>';
	strXml += ' <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
	strXml += ' <Default Extension="jpeg" ContentType="image/jpeg"/>';
	strXml += ' <Default Extension="jpg" ContentType="image/jpg"/>';

	// STEP 1: Add standard/any media types used in Presenation
	strXml += ' <Default Extension="png" ContentType="image/png"/>';
	strXml += ' <Default Extension="gif" ContentType="image/gif"/>';
	strXml += ' <Default Extension="m4v" ContentType="video/mp4"/>'; // NOTE: Hard-Code this extension as it wont be created in loop below (as extn != type)
	strXml += ' <Default Extension="mp4" ContentType="video/mp4"/>'; // NOTE: Hard-Code this extension as it wont be created in loop below (as extn != type)
	gObjPptx.slides.forEach(function(slide, idx) {
		slide.rels.forEach(function(rel, idy) {
			if (rel.type != 'image' && rel.type != 'online' && rel.type != 'chart' && rel.extn != 'm4v' && strXml.indexOf(rel.type) == -1) {
				strXml += ' <Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>';
			}
		});
	});
	strXml += ' <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
	strXml += ' <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>';

	// STEP 2: Add presentation and slide master(s)/slide(s)
	strXml += ' <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>';
	strXml += ' <Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>';
	gObjPptx.slides.forEach(function(slide, idx) {
		strXml += '<Override PartName="/ppt/slideMasters/slideMaster' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
		strXml += '<Override PartName="/ppt/slides/slide' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
		// add charts if any
		slide.rels.forEach(function(rel) {
			if (rel.type == 'chart') {
				strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
			}
		});
	});

	// STEP 3: Core PPT
	strXml += ' <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>';
	strXml += ' <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>';
	strXml += ' <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>';
	strXml += ' <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>';

	// STEP 4: Add Slide Layouts
	gObjPptx.slideLayouts.forEach(function(layout, idx) {
		strXml += '<Override PartName="/ppt/slideLayouts/slideLayout' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
		layout.rels.forEach(function(rel) {
			if (rel.type == 'chart') {
				strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
			}
		});
	});

	// STEP 5: Add notes slide(s)
	gObjPptx.slides.forEach(function(slide, idx) {
		strXml += ' <Override PartName="/ppt/notesSlides/notesSlide' + (idx + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>';
	});

	gObjPptx.masterSlide.rels.forEach(function(rel) {
		if (rel.type == 'chart') {
			strXml += ' <Override PartName="' + rel.Target + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
		}
		if (rel.type != 'image' && rel.type != 'online' && rel.type != 'chart' && rel.extn != 'm4v' && strXml.indexOf(rel.type) == -1)
			strXml += ' <Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>';
	});

	// STEP 5: Finish XML (Resume core)
	strXml += ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
	strXml += ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
	strXml += '</Types>';

	return strXml;
}

function makeXmlRootRels() {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
		+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
		+ '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
		+ '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
		+ '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
		+ '</Relationships>';
	return strXml;
}

function makeXmlApp() {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
	strXml += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
	strXml += '<TotalTime>0</TotalTime>';
	strXml += '<Words>0</Words>';
	strXml += '<Application>Microsoft Office PowerPoint</Application>';
	strXml += '<PresentationFormat>On-screen Show</PresentationFormat>';
	strXml += '<Paragraphs>0</Paragraphs>';
	strXml += '<Slides>' + gObjPptx.slides.length + '</Slides>';
	strXml += '<Notes>' + gObjPptx.slides.length + '</Notes>';
	strXml += '<HiddenSlides>0</HiddenSlides>';
	strXml += '<MMClips>0</MMClips>';
	strXml += '<ScaleCrop>false</ScaleCrop>';
	strXml += '<HeadingPairs>';
	strXml += '  <vt:vector size="4" baseType="variant">';
	strXml += '    <vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>';
	strXml += '    <vt:variant><vt:i4>1</vt:i4></vt:variant>';
	strXml += '    <vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>';
	strXml += '    <vt:variant><vt:i4>' + gObjPptx.slides.length + '</vt:i4></vt:variant>';
	strXml += '  </vt:vector>';
	strXml += '</HeadingPairs>';
	strXml += '<TitlesOfParts>';
	strXml += '<vt:vector size="' + (gObjPptx.slides.length + 1) + '" baseType="lpstr">';
	strXml += '<vt:lpstr>Office Theme</vt:lpstr>';
	gObjPptx.slides.forEach(function(slideObj, idx) { strXml += '<vt:lpstr>Slide ' + (idx + 1) + '</vt:lpstr>'; });
	strXml += '</vt:vector>';
	strXml += '</TitlesOfParts>';
	strXml += '<Company>' + gObjPptx.company + '</Company>';
	strXml += '<LinksUpToDate>false</LinksUpToDate>';
	strXml += '<SharedDoc>false</SharedDoc>';
	strXml += '<HyperlinksChanged>false</HyperlinksChanged>';
	strXml += '<AppVersion>15.0000</AppVersion>';
	strXml += '</Properties>';

	return strXml;
}

function makeXmlCore() {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
	strXml += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
	strXml += '<dc:title>' + encodeXmlEntities(gObjPptx.title) + '</dc:title>';
	strXml += '<dc:subject>' + encodeXmlEntities(gObjPptx.subject) + '</dc:subject>';
	strXml += '<dc:creator>' + encodeXmlEntities(gObjPptx.author) + '</dc:creator>';
	strXml += '<cp:lastModifiedBy>' + encodeXmlEntities(gObjPptx.author) + '</cp:lastModifiedBy>';
	strXml += '<cp:revision>' + gObjPptx.revision + '</cp:revision>';
	strXml += '<dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:created>';
	strXml += '<dcterms:modified xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:modified>';
	strXml += '</cp:coreProperties>';
	return strXml;
}

function makeXmlPresentationRels() {
	var intRelNum = 0;
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
	strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
	strXml += '  <Relationship Id="rId1" Target="slideMasters/slideMaster1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
	intRelNum++;
	for (var idx = 1; idx <= gObjPptx.slides.length; idx++) {
		intRelNum++;
		strXml += '  <Relationship Id="rId' + intRelNum + '" Target="slides/slide' + idx + '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>';
	}
	intRelNum++;
	strXml += '  <Relationship Id="rId' + intRelNum + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>'
		+ '  <Relationship Id="rId' + (intRelNum + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>'
		+ '  <Relationship Id="rId' + (intRelNum + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
		+ '  <Relationship Id="rId' + (intRelNum + 3) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>'
		+ '  <Relationship Id="rId' + (intRelNum + 4) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>'
		+ '</Relationships>';

	return strXml;
}

// XML-GEN: Next 5 functions run 1-N times (once for each Slide)

/**
 * Generates XML for the slide file
 * @param {Object} objSlide - the slide object to transform into XML
 * @return {string} strXml - slide OOXML
*/
function makeXmlSlide(objSlide) {
	// STEP 1: Generate slide XML - wrap generated text in full XML envelope
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
	strXml += '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' + (objSlide.slide.hidden ? ' show="0"' : '') + '>';
	strXml += gObjPptxGenerators.slideObjectToXml(objSlide);
	strXml += '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>';
	strXml += '</p:sld>';

	// LAST: Return
	return strXml;
}

function getNotesFromSlide(objSlide) {
	var notesStr = '';
	objSlide.data.forEach(function(data) {
		if (data.type === 'notes') {
			notesStr += data.text;
		}
	});
	return notesStr.replace(/\r*\n/g, CRLF);
}

function makeXmlNotesSlide(objSlide) {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
	strXml += '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">';
	strXml += '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="" /><p:cNvGrpSpPr />'
		+ '<p:nvPr /></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0" />'
		+ '<a:ext cx="0" cy="0" /><a:chOff x="0" y="0" /><a:chExt cx="0" cy="0" />'
		+ '</a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1" />'
		+ '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1" /></p:cNvSpPr>'
		+ '<p:nvPr><p:ph type="sldImg" /></p:nvPr></p:nvSpPr><p:spPr />'
		+ '</p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2" />'
		+ '<p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr>'
		+ '<p:ph type="body" idx="1" /></p:nvPr></p:nvSpPr><p:spPr />'
		+ '<p:txBody><a:bodyPr /><a:lstStyle /><a:p><a:r>'
		+ '<a:rPr lang="en-US" dirty="0" smtClean="0" /><a:t>'
		+ encodeXmlEntities(getNotesFromSlide(objSlide))
		+ '</a:t></a:r><a:endParaRPr lang="en-US" dirty="0" /></a:p></p:txBody>'
		+ '</p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3" />'
		+ '<p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr>'
		+ '<p:ph type="sldNum" sz="quarter" idx="10" /></p:nvPr></p:nvSpPr>'
		+ '<p:spPr /><p:txBody><a:bodyPr /><a:lstStyle /><a:p>'
		+ '<a:fld id="'
		+ SLDNUMFLDID
		+ '" type="slidenum">'
		+ '<a:rPr lang="en-US" smtClean="0" /><a:t>'
		+ objSlide.numb
		+ '</a:t></a:fld><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp>'
		+ '</p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">'
		+ '<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991" />'
		+ '</p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping /></p:clrMapOvr></p:notes>';
	return strXml;
}

/**
 * Generates the XML layout resource from a layout object
 * @param {Object} objSlideLayout - slide object that represents layout
 * @return {string} strXml - slide OOXML
*/
function makeXmlLayout(objSlideLayout) {
	// STEP 1: Generate slide XML - wrap generated text in full XML envelope
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
	strXml += '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" preserve="1">';
	strXml += gObjPptxGenerators.slideObjectToXml(objSlideLayout);
	strXml += '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>';
	strXml += '</p:sldLayout>';

	// LAST: Return
	return strXml;
}

/**
 * Generates XML for the master file
 * @param {Object} objSlide - slide object that represents master slide layout
 * @return {string} strXml - slide OOXML
*/
function makeXmlMaster(objSlide) {
	// NOTE: Pass layouts as static rels because they are not referenced any time
	var layoutDefs = gObjPptx.slideLayouts.map(function(layoutDef, idx) {
		return '<p:sldLayoutId id="' + (LAYOUT_IDX_SERIES_BASE + idx) + '" r:id="rId' + (objSlide.rels.length + idx + 1) + '"/>';
	});

	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
	strXml += '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">';
	strXml += gObjPptxGenerators.slideObjectToXml(objSlide);
	strXml += '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
	strXml += '<p:sldLayoutIdLst>' + layoutDefs.join('') + '</p:sldLayoutIdLst>';
	strXml += '<p:hf sldNum="0" hdr="0" ftr="0" dt="0"/>';
	strXml += '<p:txStyles>'
		+ ' <p:titleStyle>'
		+ '  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>'
		+ ' </p:titleStyle>'
		+ ' <p:bodyStyle>'
		+ '  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>'
		+ '  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>'
		+ '  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>'
		+ '  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>'
		+ '  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>'
		+ '  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>'
		+ '  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>'
		+ '  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>'
		+ '  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>'
		+ ' </p:bodyStyle>'
		+ ' <p:otherStyle>'
		+ '  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>'
		+ '  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>'
		+ '  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>'
		+ '  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>'
		+ '  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>'
		+ '  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>'
		+ '  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>'
		+ '  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>'
		+ '  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>'
		+ '  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>'
		+ ' </p:otherStyle>'
		+ '</p:txStyles>';
	strXml += '</p:sldMaster>';

	// LAST: Return
	return strXml;
}

function makeXmlNotesMaster() {
	return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF + '<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1" /></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="" /><p:cNvGrpSpPr /><p:nvPr /></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="0" cy="0" /><a:chOff x="0" y="0" /><a:chExt cx="0" cy="0" /></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Header Placeholder 1" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="hdr" sz="quarter" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="2971800" cy="458788" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" /><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200" /></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Date Placeholder 2" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="dt" idx="1" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="0" /><a:ext cx="2971800" cy="458788" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" /><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200" /></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{5282F153-3F37-0F45-9E97-73ACFA13230C}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0" /><a:t>6/20/18</a:t></a:fld><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Image Placeholder 3" /><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1" /></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000" /><a:ext cx="5486400" cy="3086100" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom><a:noFill /><a:ln w="12700"><a:solidFill><a:prstClr val="black" /></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr" /><a:lstStyle /><a:p><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Notes Placeholder 4" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550" /><a:ext cx="5486400" cy="3600450" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" /><a:lstStyle /><a:p><a:pPr lvl="0" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4" /><a:r><a:rPr lang="en-US" smtClean="0" /><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Footer Placeholder 5" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="4" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="8685213" /><a:ext cx="2971800" cy="458787" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b" /><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200" /></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="7" name="Slide Number Placeholder 6" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="5" /></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3884613" y="8685213" /><a:ext cx="2971800" cy="458787" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b" /><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200" /></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{CE5E9CC1-C706-0F49-92D6-E571CC5EEA8F}" type="slidenum"><a:rPr lang="en-US" smtClean="0" /><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US" /></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991" /></p:ext></p:extLst></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink" /><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1" /></a:solidFill><a:latin typeface="+mn-lt" /><a:ea typeface="+mn-ea" /><a:cs typeface="+mn-cs" /></a:defRPr></a:lvl9pPr></p:notesStyle></p:notesMaster>';
}

/**
 * Generates XML string for a slide layout relation file.
 * @param {Number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @return {String} complete XML string ready to be saved as a file
 */
function makeXmlSlideLayoutRel(layoutNumber) {
	return gObjPptxGenerators.slideObjectRelationsToXml(
		gObjPptx.slideLayouts[layoutNumber - 1],
		[{
			target: "../slideMasters/slideMaster1.xml", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
		}]
	);
}

/**
 * Generates XML string for a slide relation file.
 * @param {Number} slideNumber 1-indexed number of a layout that relations are generated for
 * @return {String} complete XML string ready to be saved as a file
 */
function makeXmlSlideRel(slideNumber) {
	return gObjPptxGenerators.slideObjectRelationsToXml(
		gObjPptx.slides[slideNumber - 1],
		[
			{
				target: '../slideLayouts/slideLayout' + getLayoutIdxForSlide(slideNumber) + '.xml',
				type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
			},
			{
				target: '../notesSlides/notesSlide' + slideNumber + '.xml',
				type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
			}
		]
	);
}

function makeXmlNotesSlideRel(slideNumber) {
	return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
		+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
		+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>'
		+ '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide' + slideNumber + '.xml"/>'
		+ '</Relationships>';
}

/**
 * Generates XML string for the master file.
 * @param {Object} masterSlideObject slide object
 * @return {String} complete XML string ready to be saved as a file
 */
function makeXmlMasterRel(masterSlideObject) {
	var relCount = masterSlideObject.rels.length
	var defaultRels = gObjPptx.slideLayouts.map(function(layoutDef, idx) {
		return { target: '../slideLayouts/slideLayout' + (idx + 1) + '.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' };
	});
	defaultRels.push({ target: '../theme/theme1.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' });

	return gObjPptxGenerators.slideObjectRelationsToXml(
		masterSlideObject,
		defaultRels
	);
}

function makeXmlNotesMasterRel() {
	return '<?xml version="1.0" encoding="UTF-8"?>' + CRLF
		+ '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
		+ '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
		+ '</Relationships>';
}

/**
 * For the passed slide number, resolves name of a layout that is used for.
 * @param {Number} slideNumber
 * @return {Number} slide number
 */
function getLayoutIdxForSlide(slideNumber) {
	var layoutName = gObjPptx.slides[slideNumber - 1].layout;
	var layoutIdx = -1;

	for (var i = 0; i < gObjPptx.slideLayouts.length; i++) {
		if (gObjPptx.slideLayouts[i].name === layoutName) {
			return (i + 1);
		}
	}

	// IMPORTANT: Return 1 (for `slideLayout1.xml`) when no def is found
	// So all objects are in Layout1 and every slide that references it uses this layout.
	return 1;
}

// XML-GEN: Last 5 functions create root /ppt files

function makeXmlTheme() {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF;
	strXml += '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">\
					<a:themeElements>\
					  <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>\
					  <a:dk2><a:srgbClr val="A7A7A7"/></a:dk2>\
					  <a:lt2><a:srgbClr val="535353"/></a:lt2>\
					  <a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>\
					  <a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>\
					  <a:fontScheme name="Office">\
					  <a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont>\
					  <a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/>\
					  </a:minorFont></a:fontScheme>\
					  <a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>\
					</a:theme>';
	return strXml;
}

function makeXmlPresentation() {
	var intCurPos = 0;
	// REF: http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
		+ '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
		+ 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
		+ 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' + (gObjPptx.rtlMode ? 'rtl="1"' : '') + ' saveSubsetFonts="1" autoCompressPictures="0">';

	// STEP 1: Build SLIDE master list
	strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>';
	strXml += '<p:sldIdLst>';
	for (var idx = 0; idx < gObjPptx.slides.length; idx++) {
		strXml += '<p:sldId id="' + (idx + 256) + '" r:id="rId' + (idx + 2) + '"/>';
	}
	strXml += '</p:sldIdLst>';

	// Step 2: Add NOTES master list
	strXml += '<p:notesMasterIdLst><p:notesMasterId r:id="rId' + (gObjPptx.slides.length + 2 + 4) + '"/></p:notesMasterIdLst>'; // length+2+4 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic)

	// STEP 3: Build SLIDE text styles
	strXml += '<p:sldSz cx="' + gObjPptx.pptLayout.width + '" cy="' + gObjPptx.pptLayout.height + '" type="' + gObjPptx.pptLayout.name + '"/>'
		+ '<p:notesSz cx="' + gObjPptx.pptLayout.height + '" cy="' + gObjPptx.pptLayout.width + '"/>'
		+ '<p:defaultTextStyle>';
	+ '  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>';
	for (var idx = 1; idx < 10; idx++) {
		strXml += '  <a:lvl' + idx + 'pPr marL="' + intCurPos + '" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
			+ '    <a:defRPr sz="1800" kern="1200">'
			+ '      <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
			+ '      <a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>'
			+ '    </a:defRPr>'
			+ '  </a:lvl' + idx + 'pPr>';
		intCurPos += 457200;
	}
	strXml += '</p:defaultTextStyle>';
	strXml += '</p:presentation>';

	return strXml;
}

function makeXmlPresProps() {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
		+ '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'

	return strXml;
}

function makeXmlTableStyles() {
	// SEE: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
		+ '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>';
	return strXml;
}

function makeXmlViewProps() {
	var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
		+ '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
		+ '<p:normalViewPr><p:restoredLeft sz="15610" /><p:restoredTop sz="94613" /></p:normalViewPr>'
		+ '<p:slideViewPr>'
		+ '  <p:cSldViewPr snapToGrid="0" snapToObjects="1">'
		+ '    <p:cViewPr varScale="1"><p:scale><a:sx n="119" d="100" /><a:sy n="119" d="100" /></p:scale><p:origin x="312" y="184" /></p:cViewPr>'
		+ '    <p:guideLst />'
		+ '  </p:cSldViewPr>'
		+ '</p:slideViewPr>'
		+ '<p:notesTextViewPr>'
		+ '  <p:cViewPr><p:scale><a:sx n="1" d="1" /><a:sy n="1" d="1" /></p:scale><p:origin x="0" y="0" /></p:cViewPr>'
		+ '</p:notesTextViewPr>'
		+ '<p:gridSpacing cx="76200" cy="76200" />'
		+ '</p:viewPr>';
	return strXml;
}

/**
 * @param {Object} glOpts {size, color, style}
 * @param {Object} defaults {size, color, style}
 * @param {String} type "major"(default) | "minor"
 */
function createGridLineElement(glOpts: { size: number, color: string, style: string }, defaults: { size: number, color: string, style: string }, type: "major" | "minor") {
	type = type || 'major';
	let tagName = 'c:' + type + 'Gridlines';
	let strXml = '<' + tagName + '>';
	strXml += ' <c:spPr>';
	strXml += '  <a:ln w="' + Math.round((glOpts.size || defaults.size) * ONEPT) + '" cap="flat">';
	strXml += '  <a:solidFill><a:srgbClr val="' + (glOpts.color || defaults.color) + '"/></a:solidFill>'; // should accept scheme colors as implemented in PR 135
	strXml += '   <a:prstDash val="' + (glOpts.style || defaults.style) + '"/><a:round/>';
	strXml += '  </a:ln>';
	strXml += ' </c:spPr>';
	strXml += '</' + tagName + '>';
	return strXml;
}

/**
 * DESC: Calc and return excel column name (eg: 'A2')
 */
function getExcelColName(length) {
	var strName = '';

	if (length <= 26) {
		strName = LETTERS[length];
	}
	else {
		strName += LETTERS[Math.floor(length / LETTERS.length) - 1];
		strName += LETTERS[(length % LETTERS.length)];
	}

	return strName;
}

/**
 * DESC: Depending on the passed color string, creates either `a:schemeClr` (when scheme color) or `a:srgbClr` (when hexa representation).
 * color (string): hexa representation (eg. "FFFF00") or a scheme color constant (eg. colors.ACCENT1)
 * innerElements (optional string): Additional elements that adjust the color and are enclosed by the color element.
 */
function createColorElement(colorStr, innerElements) {
	var isHexaRgb = REGEX_HEX_COLOR.test(colorStr);
	if (!isHexaRgb && SCHEME_COLOR_NAMES.indexOf(colorStr) === -1) {
		console.warn('"' + colorStr + '" is not a valid scheme color or hexa RGB! "' + DEF_FONT_COLOR + '" is used as a fallback. Pass 6-digit RGB or these `pptx.colors` values:\n' + SCHEME_COLOR_NAMES.join(', '));
		colorStr = DEF_FONT_COLOR;
	}
	var tagName = isHexaRgb ? 'srgbClr' : 'schemeClr';
	var colorAttr = ' val="' + colorStr + '"';
	return innerElements ? '<a:' + tagName + colorAttr + '>' + innerElements + '</a:' + tagName + '>' : '<a:' + tagName + colorAttr + ' />';
}
