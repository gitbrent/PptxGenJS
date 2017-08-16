var DEF_FONT_SIZE = 12;
var LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
var DEF_CHART_GRIDLINE = { color: "888888", style: "solid", size: 1 };
var ONEPT = 12700; // One (1) point (pt)
var PIECHART_COLORS = ['5DA5DA','FAA43A','60BD68','F17CB0','B2912F','B276B2','DECF3F','F15854','A7A7A7', '5DA5DA','FAA43A','60BD68','F17CB0','B2912F','B276B2','DECF3F','F15854','A7A7A7'];
var BARCHART_COLORS = ['C0504D','4F81BD','9BBB59','8064A2','4BACC6','F79646','628FC6','C86360', 'C0504D','4F81BD','9BBB59','8064A2','4BACC6','F79646','628FC6','C86360'];


function createGridLineElement(glOpts, defaults, type) {
	type = type || 'major';
	var tagName = 'c:'+ type + 'Gridlines';
	strXml =  '<'+ tagName + '>';
	strXml += ' <c:spPr>';
	strXml += '  <a:ln w="' + Math.round((glOpts.size || defaults.size) * ONEPT) +'" cap="flat">';
	strXml += '  <a:solidFill><a:srgbClr val="' + (glOpts.color || defaults.color) + '"/></a:solidFill>'; // should accept scheme colors as implemented in PR 135
	strXml += '   <a:prstDash val="' + (glOpts.style || defaults.style) + '"/><a:round/>';
	strXml += '  </a:ln>';
	strXml += ' </c:spPr>';
	strXml += '</'+ tagName + '>';
	return strXml;
}

function makeChartType (chartType, data, opts) {

	console.log('makeChartType', chartType, data, opts);

	function getExcelColName(length) {
		var strName = '';

		if ( length <= 26 ) {
			strName = LETTERS[length];
		}
		else {
			strName += LETTERS[ Math.floor(length/LETTERS.length)-1 ];
			strName += LETTERS[ (length % LETTERS.length) ];
		}

		return strName;
	}

	var strXml = '';
	switch ( chartType ) {
		case 'area':
		case 'bar':
		case 'line':
			strXml += '<c:'+ chartType +'Chart>';
			if ( chartType == 'bar' ) strXml += '  <c:barDir val="'+ opts.barDir +'"/>';
			strXml += '  <c:grouping val="'+ opts.barGrouping + '"/>';
			strXml += '  <c:varyColors val="0"/>';

			// A: "Series" block for every data row
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

			// this needs to maintain the index depending on the region.... how????????????

			data.forEach(function(obj){
				var idx = obj.index;
				console.log('data', idx, obj);
				strXml += '<c:ser>';
				strXml += '  <c:idx val="'+ idx +'"/>';
				strXml += '  <c:order val="'+ idx +'"/>';
				strXml += '  <c:tx>';
				strXml += '    <c:strRef>';
				strXml += '      <c:f>Sheet1!$A$'+ (idx+2) +'</c:f>';
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>'+ obj.name +'</c:v></c:pt></c:strCache>';
				strXml += '    </c:strRef>';
				strXml += '  </c:tx>';

				// Fill and Border
				var strSerColor = opts.chartColors[(idx+1 > opts.chartColors.length ? (Math.floor(Math.random() * opts.chartColors.length)) : idx)];
				strXml += '  <c:spPr>';

				if ( opts.chartColorsOpacity ) {
					strXml += '    <a:solidFill><a:srgbClr val="'+ strSerColor +'"><a:alpha val="50000"/></a:srgbClr></a:solidFill>';
				}
				else {
					strXml += '    <a:solidFill><a:srgbClr val="'+ strSerColor +'"/></a:solidFill>';
				}

				if ( chartType == 'line' ) {
					strXml += '<a:ln w="'+ (opts.lineSize * ONEPT) +'" cap="flat"><a:solidFill><a:srgbClr val="'+ strSerColor +'"/></a:solidFill>';
					strXml += '<a:prstDash val="' + (opts.line_dash || "solid") + '"/><a:round/></a:ln>';
				}
				else if ( opts.dataBorder ) {
					strXml += '<a:ln w="'+ (opts.dataBorder.pt * ONEPT) +'" cap="flat"><a:solidFill><a:srgbClr val="'+ opts.dataBorder.color +'"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
				}
				strXml += '    <a:effectLst>';
				strXml += '      <a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="tl" rotWithShape="1" blurRad="38100" dist="23000" dir="5400000">';
				strXml += '        <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>';
				strXml += '      </a:outerShdw>';
				strXml += '    </a:effectLst>';
				strXml += '  </c:spPr>';

				// LINE CHART ONLY: `marker`
				if ( chartType == 'line' ) {
					strXml += '<c:marker>';
					strXml += '  <c:symbol val="'+ opts.lineDataSymbol +'"/>';
					if ( opts.lineDataSymbolSize ) strXml += '  <c:size val="'+ opts.lineDataSymbolSize +'"/>'; // Defaults to "auto" otherwise (but this is usually too small, so there is a default)
					strXml += '  <c:spPr>';
					strXml += '    <a:solidFill><a:srgbClr val="'+ opts.chartColors[(idx+1 > opts.chartColors.length ? (Math.floor(Math.random() * opts.chartColors.length)) : idx)] +'"/></a:solidFill>';
					strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="'+ strSerColor +'"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
					strXml += '    <a:effectLst/>';
					strXml += '  </c:spPr>';
					strXml += '</c:marker>';
				}

				// Color bar chart bars various colors
				// Allow users with a single data set to pass their own array of colors (check for this using != ours)
				if ( data.length === 1 && opts.chartColors != BARCHART_COLORS ) {
					// Series Data Point colors
					obj.values.forEach(function(value,index){
						strXml += '  <c:dPt>';
						strXml += '    <c:idx val="'+index+'"/>';
						strXml += '    <c:invertIfNegative val="1"/>';
						strXml += '    <c:bubble3D val="0"/>';
						strXml += '    <c:spPr>';
						strXml += '    <a:solidFill>';
						strXml += '     <a:srgbClr val="'+opts.chartColors[index % opts.chartColors.length]+'"/>';
						strXml += '    </a:solidFill>';
						strXml += '    <a:effectLst>';
						strXml += '    <a:outerShdw blurRad="38100" dist="23000" dir="5400000" algn="tl">';
						strXml += '    	<a:srgbClr val="000000">';
						strXml += '    	<a:alpha val="35000"/>';
						strXml += '    	</a:srgbClr>';
						strXml += '    </a:outerShdw>';
						strXml += '    </a:effectLst>';
						strXml += '    </c:spPr>';
						strXml += '  </c:dPt>';
					});
				}



				// 2: "Categories"
				{
					strXml += '<c:cat>';
					strXml += '  <c:strRef>';
					strXml += '    <c:f>Sheet1!' + '$B$1:$' + getExcelColName(obj.labels.length) + '$1' + '</c:f>';
					strXml += '    <c:strCache>';
					strXml += '	     <c:ptCount val="' + obj.labels.length + '"/>';
					obj.labels.forEach(function (label, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + label + '</c:v></c:pt>'; });
					strXml += '    </c:strCache>';
					strXml += '  </c:strRef>';
					strXml += '</c:cat>';
				}

				// 3: "Values"
				{
					strXml += '  <c:val>';
					strXml += '    <c:numRef>';
					strXml += '      <c:f>Sheet1!' + '$B$' + (idx + 2) + ':$' + getExcelColName(obj.labels.length) + '$' + (idx + 2) + '</c:f>';
					strXml += '      <c:numCache>';
					strXml += '	       <c:ptCount val="' + obj.labels.length + '"/>';
					obj.values.forEach(function (value, idx) { strXml += '<c:pt idx="' + idx + '"><c:v>' + value + '</c:v></c:pt>'; });
					strXml += '      </c:numCache>';
					strXml += '    </c:numRef>';
					strXml += '  </c:val>';
				}

				// LINE CHART ONLY: `smooth`
				if ( chartType == 'line' ) strXml += '<c:smooth val="'+ (opts.lineSmooth ? "1" : "0" ) +'"/>';

				// 4: Close "SERIES"
				strXml += '</c:ser>';

			}); // end forEach

			// 1: "Data Labels"
			{
				strXml += '  <c:dLbls>';
				strXml += '    <c:numFmt formatCode="' + opts.dataLabelFormatCode + '" sourceLinked="0"/>';
				strXml += '    <c:txPr>';
				strXml += '      <a:bodyPr/>';
				strXml += '      <a:lstStyle/>';
				strXml += '      <a:p><a:pPr>';
				strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="' + (opts.dataLabelFontSize || DEF_FONT_SIZE) + '00" u="none">';
				strXml += '          <a:solidFill><a:srgbClr val="' + (opts.dataLabelColor || '000000') + '"/></a:solidFill>';
				strXml += '          <a:latin typeface="' + (opts.dataLabelFontFace || 'Arial') + '"/>';
				strXml += '        </a:defRPr>';
				strXml += '      </a:pPr></a:p>';
				strXml += '    </c:txPr>';
				if (chartType != 'area') strXml += '    <c:dLblPos val="' + (opts.dataLabelPosition || 'outEnd') + '"/>';
				strXml += '    <c:showLegendKey val="0"/>';
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>';
				strXml += '    <c:showCatName val="0"/>';
				strXml += '    <c:showSerName val="0"/>';
				strXml += '    <c:showPercent val="0"/>';
				strXml += '    <c:showBubbleSize val="0"/>';
				strXml += '    <c:showLeaderLines val="0"/>';
				strXml += '  </c:dLbls>';
			}

			if ( chartType == 'bar' ) {
				strXml += '  <c:gapWidth val="'+ opts.barGapWidthPct +'"/>';
				strXml += '  <c:overlap val="'+ (opts.barGrouping.indexOf('tacked') > -1 ? 100 : 0) +'"/>';
			}
			else if ( chartType == 'line' ) {
				strXml += '  <c:marker val="1"/>';
			}
			strXml += '  <c:axId val="2094734552"/>';
			strXml += '  <c:axId val="2094734553"/>';
			strXml += '</c:'+ chartType +'Chart>';

			// Done with CHART.BAR/LINE
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

			// 1: Start pieChart
			strXml += '<c:'+ chartType +'Chart>';
			strXml += '  <c:varyColors val="0"/>';
			strXml += '<c:ser>';
			strXml += '  <c:idx val="0"/>';
			strXml += '  <c:order val="0"/>';
			strXml += '  <c:tx>';
			strXml += '    <c:strRef>';
			strXml += '      <c:f>Sheet1!$A$2</c:f>';
			strXml += '      <c:strCache>';
			strXml += '        <c:ptCount val="1"/>';
			strXml += '        <c:pt idx="0"><c:v>'+ decodeXmlEntities(obj.name) +'</c:v></c:pt>';
			strXml += '      </c:strCache>';
			strXml += '    </c:strRef>';
			strXml += '  </c:tx>';
			strXml += '  <c:spPr>';
			strXml += '    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
			strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
			strXml += '    <a:effectLst>';
			strXml += '      <a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="tl" rotWithShape="1" blurRad="38100" dist="23000" dir="5400000">';
			strXml += '        <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>';
			strXml += '      </a:outerShdw>';
			strXml += '    </a:effectLst>';
			strXml += '  </c:spPr>';
			strXml += '<c:explosion val="0"/>';

			// 2: "Data Point" block for every data row
			obj.labels.forEach(function(label,idx){
				strXml += '<c:dPt>';
				strXml += '  <c:idx val="'+ idx +'"/>';
				strXml += '  <c:explosion val="0"/>';
				strXml += '  <c:spPr>';
				strXml += '    <a:solidFill><a:srgbClr val="'+ opts.chartColors[(idx+1 > opts.chartColors.length ? (Math.floor(Math.random() * opts.chartColors.length)) : idx)] +'"/></a:solidFill>';
				if ( opts.dataBorder ) {
					strXml += '<a:ln w="'+ (opts.dataBorder.pt * ONEPT) +'" cap="flat"><a:solidFill><a:srgbClr val="'+ opts.dataBorder.color +'"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
				}
				strXml += '    <a:effectLst>';
				strXml += '      <a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="tl" rotWithShape="1" blurRad="38100" dist="23000" dir="5400000">';
				strXml += '        <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>';
				strXml += '      </a:outerShdw>';
				strXml += '    </a:effectLst>';
				strXml += '  </c:spPr>';
				strXml += '</c:dPt>';
			});

			// 3: "Data Label" block for every data Label
			strXml += '<c:dLbls>';
			obj.labels.forEach(function(label,idx){
				strXml += '<c:dLbl>';
				strXml += '  <c:idx val="'+ idx +'"/>';
				strXml += '    <c:numFmt formatCode="'+ opts.dataLabelFormatCode +'" sourceLinked="0"/>';
				strXml += '    <c:txPr>';
				strXml += '      <a:bodyPr/><a:lstStyle/>';
				strXml += '      <a:p><a:pPr>';
				strXml += '        <a:defRPr b="0" i="0" strike="noStrike" sz="'+ (opts.dataLabelFontSize || DEF_FONT_SIZE) +'00" u="none">';
				strXml += '          <a:solidFill><a:srgbClr val="'+ (opts.dataLabelColor || '000000') +'"/></a:solidFill>';
				strXml += '          <a:latin typeface="'+ (opts.dataLabelFontFace || 'Arial') +'"/>';
				strXml += '        </a:defRPr>';
				strXml += '      </a:pPr></a:p>';
				strXml += '    </c:txPr>';
				if (chartType == 'pie') {
					strXml += '    <c:dLblPos val="'+ (opts.dataLabelPosition || 'inEnd') +'"/>';
				}
				strXml += '    <c:showLegendKey val="0"/>';
				strXml += '    <c:showVal val="'+ (opts.showValue ? "1" : "0") +'"/>';
				strXml += '    <c:showCatName val="'+ (opts.showLabel ? "1" : "0") +'"/>';
				strXml += '    <c:showSerName val="0"/>';
				strXml += '    <c:showPercent val="'+ (opts.showPercent ? "1" : "0") +'"/>';
				strXml += '    <c:showBubbleSize val="0"/>';
				strXml += '  </c:dLbl>';
			});
			strXml += '<c:numFmt formatCode="'+ opts.dataLabelFormatCode +'" sourceLinked="0"/>\
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
			strXml += '    <c:f>Sheet1!'+ '$B$1:$'+ getExcelColName(obj.labels.length) +'$1' +'</c:f>';
			strXml += '    <c:strCache>';
			strXml += '	     <c:ptCount val="'+ obj.labels.length +'"/>';
			obj.labels.forEach(function(label,idx){ strXml += '<c:pt idx="'+ idx +'"><c:v>'+ label +'</c:v></c:pt>'; });
			strXml += '    </c:strCache>';
			strXml += '  </c:strRef>';
			strXml += '</c:cat>';

			// 3: Create vals
			strXml += '  <c:val>';
			strXml += '    <c:numRef>';
			strXml += '      <c:f>Sheet1!'+ '$B$2:$'+ getExcelColName(obj.labels.length) +'$'+ 2 +'</c:f>';
			strXml += '      <c:numCache>';
			strXml += '	       <c:ptCount val="'+ obj.labels.length +'"/>';
			obj.values.forEach(function(value,idx){ strXml += '<c:pt idx="'+ idx +'"><c:v>'+ value +'</c:v></c:pt>'; });
			strXml += '      </c:numCache>';
			strXml += '    </c:numRef>';
			strXml += '  </c:val>';

			// 4: Close "SERIES"
			strXml += '  </c:ser>';
			strXml += '  <c:firstSliceAng val="0"/>';
			if ( chartType == 'doughnut' ) strXml += '  <c:holeSize val="' + (opts.holeSize || 50) + '"/>';
			strXml += '</c:'+ chartType +'Chart>';

			// Done with CHART.PIE
			break;
	}

	return strXml;
}

function makeChartAxes (chartType, opts) {
	console.log('makeChartAxes', chartType);
	var strXml = '';

	if(chartType === 'pie' || chartType === 'doughnut'){
		return strXml;
	}
	// B: "Category Axis"
	{
		strXml += '<c:catAx>';
		if (opts.showCatAxisTitle) {
			strXml += genXmlTitle({
				title: opts.catAxisTitle || 'Axis Title',
				fontSize: opts.catAxisTitleFontSize,
				color: opts.catAxisTitleColor,
				fontFace: opts.catAxisTitleFontFace,
				rotate: opts.catAxisTitleRotate
			});
		}
		strXml += '  <c:axId val="2094734552"/>';
		strXml += '  <c:scaling><c:orientation val="'+ (opts.catAxisOrientation || (opts.barDir == 'col' ? 'minMax' : 'minMax')) +'"/></c:scaling>';
		strXml += '  <c:delete val="'+ (opts.catAxisHidden ? 1 : 0) +'"/>';
		strXml += '  <c:axPos val="'+ (opts.barDir == 'col' ? 'b' : 'l') +'"/>';
		if ( opts.catGridLine !== 'none' ) {
			strXml += createGridLineElement(opts.catGridLine, DEF_CHART_GRIDLINE);
		}
		strXml += '  <c:numFmt formatCode="General" sourceLinked="0"/>';
		strXml += '  <c:majorTickMark val="out"/>';
		strXml += '  <c:minorTickMark val="none"/>';
		strXml += '  <c:tickLblPos val="'+ (opts.barDir == 'col' ? 'low' : 'nextTo') +'"/>';
		strXml += '  <c:spPr>';
		strXml += '    <a:ln w="12700" cap="flat"><a:solidFill><a:srgbClr val="888888"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
		strXml += '  </c:spPr>';
		strXml += '  <c:txPr>';
		strXml += '    <a:bodyPr rot="0"/>';
		strXml += '    <a:lstStyle/>';
		strXml += '    <a:p>';
		strXml += '    <a:pPr>';
		strXml += '<a:defRPr b="0" i="0" strike="noStrike" sz="'+ (opts.catAxisLabelFontSize || DEF_FONT_SIZE) +'00" u="none">';
		strXml += '<a:solidFill><a:srgbClr val="'+ (opts.catAxisLabelColor || '000000') +'"/></a:solidFill>';
		strXml += '<a:latin typeface="'+ (opts.catAxisLabelFontFace || 'Arial') +'"/>';
		strXml += '   </a:defRPr>';
		strXml += '  </a:pPr>';
		strXml += '  </a:p>';
		strXml += ' </c:txPr>';
		strXml += ' <c:crossAx val="2094734553"/>';
		strXml += ' <c:crosses val="autoZero"/>';
		strXml += ' <c:auto val="1"/>';
		strXml += ' <c:lblAlgn val="ctr"/>';
		strXml += ' <c:noMultiLvlLbl val="1"/>';
		strXml += '</c:catAx>';
	}

	// C: "Value Axis"
	{
		strXml += '<c:valAx>';
		if (opts.showValAxisTitle) {
			strXml += genXmlTitle({
				title: opts.valAxisTitle || 'Axis Title',
				fontSize: opts.valAxisTitleFontSize,
				color: opts.valAxisTitleColor,
				fontFace: opts.valAxisTitleFontFace,
				rotate: opts.valAxisTitleRotate
			});
		}
		strXml += '  <c:axId val="2094734553"/>';
		strXml += '  <c:scaling>';
		strXml += '    <c:orientation val="'+ (opts.valAxisOrientation || (opts.barDir == 'col' ? 'minMax' : 'minMax')) +'"/>';
		if (opts.valAxisMaxVal) strXml += '<c:max val="'+ opts.valAxisMaxVal +'"/>';
		if (opts.valAxisMinVal) strXml += '<c:min val="'+ opts.valAxisMinVal +'"/>';
		strXml += '  </c:scaling>';
		strXml += '  <c:delete val="'+ (opts.valAxisHidden ? 1 : 0) +'"/>';
		strXml += '  <c:axPos val="'+ (opts.barDir == 'col' ? 'l' : 'b') +'"/>';
		if (opts.valGridLine != 'none') strXml += createGridLineElement(opts.valGridLine, DEF_CHART_GRIDLINE);
		strXml += ' <c:numFmt formatCode="'+ (opts.valAxisLabelFormatCode ? opts.valAxisLabelFormatCode : 'General') +'" sourceLinked="0"/>';
		strXml += ' <c:majorTickMark val="out"/>';
		strXml += ' <c:minorTickMark val="none"/>';
		strXml += ' <c:tickLblPos val="'+ (opts.barDir == 'col' ? 'nextTo' : 'low') +'"/>';
		strXml += ' <c:spPr>';
		strXml += '  <a:ln w="12700" cap="flat"><a:solidFill><a:srgbClr val="888888"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>';
		strXml += ' </c:spPr>';
		strXml += ' <c:txPr>';
		strXml += '  <a:bodyPr rot="0"/>';
		strXml += '  <a:lstStyle/>';
		strXml += '  <a:p>';
		strXml += '    <a:pPr>';
		strXml += '      <a:defRPr b="0" i="0" strike="noStrike" sz="'+ (opts.valAxisLabelFontSize || DEF_FONT_SIZE) +'00" u="none">';
		strXml += '        <a:solidFill><a:srgbClr val="'+ (opts.valAxisLabelColor || '000000') +'"/></a:solidFill>';
		strXml += '        <a:latin typeface="'+ (opts.valAxisLabelFontFace || 'Arial') +'"/>';
		strXml += '      </a:defRPr>';
		strXml += '    </a:pPr>';
		strXml += '  </a:p>';
		strXml += ' </c:txPr>';
		strXml += ' <c:crossAx val="2094734552"/>';
		strXml += ' <c:crosses val="autoZero"/>';
		strXml += ' <c:crossBetween val="'+ ( chartType == 'area' ? 'midCat' : 'between' ) +'"/>';
		if ( opts.valAxisMajorUnit ) strXml += ' <c:majorUnit val="'+ opts.valAxisMajorUnit +'"/>';
		strXml += '</c:valAx>';
	}

	return strXml;
}