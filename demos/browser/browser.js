function runAllDemos() {
	if (console.time) console.time('runAllDemos');
	$('#modalBusy').modal('show');

	runEveryTest()
	.catch(function(err) {
		console.error(err.toString());
		$('#modalBusy').modal('hide');
	})
	.then(function() {
		if (console.timeEnd) console.timeEnd('runAllDemos');
		$('#modalBusy').modal('hide');
	});
}

function execGenSlidesFunc(type) {
	if (console.time) console.time('execGenSlidesFunc: '+type);
	$('#modalBusy').modal('show');

	execGenSlidesFuncs(type)
	.catch(function(err) {
		$('#modalBusy').modal('hide');
		console.error(err);
	})
	.then(function() {
		$('#modalBusy').modal('hide');
		if (console.timeEnd) console.timeEnd('execGenSlidesFunc: '+type);
	});
}

/* DESC: old/undocumented/unused */
function doTestSimple() {
	var pptx = new PptxGenJS();
	var slide = pptx.addSlide();
	var optsTitle = { color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] };

	pptx.layout({ name:'A3', width:16.5, height:11.7 });
	slide.slideNumber({ x:0.5, y:'90%' });
	slide.addTable( [ [{ text:'Simple Example', options:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

	//slide.addText('Hello World!', { x:0.5, y:0.7, w:6, h:1, color:'0000FF' });
	slide.addText('Hello 45! ', { x:0.5, y:0.5, w:6, h:1, fontSize:36, color:'0000FF', shadow:{type:'outer', color:'00AAFF', blur:2, offset:10, angle: 45, opacity:0.25} });
	slide.addText('Hello 180!', { x:0.5, y:1.0, w:6, h:1, fontSize:36, color:'0000FF', shadow:{type:'outer', color:'ceAA00', blur:2, offset:10, angle:180, opacity:0.5} });
	slide.addText('Hello 355!', { x:0.5, y:1.5, w:6, h:1, fontSize:36, color:'0000FF', shadow:{type:'outer', color:'aaAA33', blur:2, offset:10, angle:355, opacity:0.75} });

	// Bullet Test: Number
	slide.addText(999, { x:0.5, y:2.0, w:'50%', h:1, color:'0000DE', bullet:true });
	// Bullet Test: Text test
	slide.addText('Bullet text', { x:0.5, y:2.5, w:'50%', h:1, color:'00AA00', bullet:true });
	// Bullet Test: Multi-line text test
	slide.addText('Line 1\nLine 2\nLine 3', { x:0.5, y:3.5, w:'50%', h:1, color:'AACD00', bullet:true });

	// Table cell margin:0
	slide.addTable([['margin:0']], { x: 0.5, y: 1.1, margin: 0, w: 0.75, fill: { color: 'FFFCCC' } });

	// Fine-grained Formatting/word-level/line-level Formatting
	slide.addText(
		[
			{ text:'right line', options:{ fontSize:24, fontFace:'Courier New', color:'99ABCC', align:'right', breakLine:true } },
			{ text:'ctr line',   options:{ fontSize:36, fontFace:'Arial',       color:'FFFF00', align:'center', breakLine:true } },
			{ text:'left line',  options:{ fontSize:48, fontFace:'Verdana',     color:'0088CC', align:'left' } }
		],
		{ x: 0.5, y: 3.0, w: 8.5, h: 4, margin: 0.1, fill: { color: '232323' } }
	);

	// Export:
	pptx.writeFile({ fileName: 'Sample Presentation' });
}

/* The "Text" demo on the PptxGenJS homepage - codified here so we can quickly reproduce the screencaps, etc. as needed */
function doHomepageDemo_Text() {
	var pptx = new PptxGenJS();
	pptx.layout ='LAYOUT_WIDE';
	var slide = pptx.addSlide();

	slide.addText(
		'BONJOUR - CIAO - GUTEN TAG - HELLO - HOLA - \nNAMASTE - OLÀ - ZDRAS-TVUY-TE - こんにちは - 你好',
		{ x:0.0, y:0.0, w:'100%', h:1.25, align:'center', fontSize:18, color:'0088CC', fill:{color:'F1F1F1'} }
	);

	slide.addText("Line-Level Formatting:", { x:0.5, y:1.5, w:'40%', h:0.38, color:'0088CC' });
	slide.addText(
		[
			{ text:'1st line', options:{ fontSize:24, fontFace:'Courier New', color:'99ABCC', align:'right', breakLine:true } },
			{ text:'2nd line', options:{ fontSize:36, fontFace:'Arial',       color:'FFFF00', align:'center', breakLine:true } },
			{ text:'3rd line', options:{ fontSize:48, fontFace:'Verdana',     color:'0088CC', align:'left' } }
		],
		{ x:0.5, y:2.0, w:6, h:2.25, margin:0.1, fill:{color:'232323'} }
	);

	slide.addText("Bullets: Normal", { x:8.0, y:1.5, w:'40%', h:0.38, color:'0088CC' });
	slide.addText(
		'Line 1\nLine 2\nLine 3',
		{ x:8.0, y:2.0, w:'30%', h:1, color:'393939', fontSize:16, fill:{color:'F2F2F2'}, bullet:true }
	);

	slide.addText("Bullets: Numbered", { x:8.0, y:3.4, w:'40%', h:0.38, color:'0088CC' });
	slide.addText(
		'Line 1\nLine 2\nLine 3',
		{ x:8.0, y:3.9, w:'30%', h:1, color:'393939', fontSize:16, fill:{color:'F2F2F2'}, bullet:{type:'number'} }
	);

	slide.addText("Bullets: Custom", { x:8.0, y:5.3, w:'40%', h:0.38, color:'0088CC' });
	slide.addText('Star bullet! ',   { x:8.0, y:5.6, w:'40%', h:0.38, color:'CC0000', bullet:{code:'2605'} });
	slide.addText('Check bullet!',   { x:8.0, y:5.9, w:'40%', h:0.38, color:'00CD00', bullet:{code:'2713'} });

	var shadowOpts = { type:'outer', color:'696969', blur:3, offset:10, angle:45, opacity:0.8 };
	slide.addText("Text Shadow:", { x:0.5, y:6.0, w:'40%', h:0.38, color:'0088CC' });
	slide.addText(
		'Outer Shadow (blur:3, offset:10, angle:45, opacity:80%)',
		{ x:0.5, y:6.4, w:12, h:0.6, fontSize:32, color:'0088cc', shadow:shadowOpts }
	);

	pptx.writeFile({ fileName: 'Demo-Text' });
}

// UNDOCUMENTED: Run from console
function testTTS() {
	var pptx = new PptxGenJS();
	pptx.layout = 'LAYOUT_WIDE';
	/*
	var slide = pptx.addSlide();
	slide.addText('Table Paging Logic Check', { x:0.0, y:'90%', w:'100%', align:'center', fontSize:18, color:'0088CC', fill:{color:'F2F9FC'} });
	var numMargin = 1.25;
	slide.addShape(pptx.shapes.RECTANGLE, { x:0.0, y:0.0, w:numMargin, h:numMargin, fill:{color:'FFFCCC'} });
	slide.addTable( ['short','table','whatever'], {x:numMargin, y:numMargin, margin:numMargin, colW:2.5, fill:{color:'F1F1F1'}} );
	*/

	// Mimic "table2slides1()"
	// addMasterDefs(pptx);

	// TEST
	//pptx.tableToSlides('tabAutoPaging');
	pptx.tableToSlides('tabAutoPaging', {verbose:true, autoPageRepeatHeader:true /*, autoPageSlideStartY:2*/});

	pptx.writeFile({ fileName: 'PptxGenJs_TTSTest_'+getTimestamp() });
}

// UNDOCUMENTED:
function testTTSMulti() {
	var ttsTitleText = { fontSize:14, color:'0088CC', bold:true };
	var ttsMultiOpts = { fontSize:13, color:'9F9F9F', verbose:true };
	var arrRows = [];
	var arrText = [];
	//
	var pptx = new PptxGenJS();
	pptx.layout = 'LAYOUT_WIDE';

	for (var idx=0; idx<gArrNamesF.length; idx++) {
		var strText = ( idx == 0 ? gStrLoremIpsum.substring(0,100) : gStrLoremIpsum.substring(idx*100,idx*200) );
		arrRows.push( [idx, gArrNamesF[idx], strText] );
		arrText.push( [strText] );
	}

	// autoPageLineWeight option demos
	var slide = pptx.addSlide();
	slide.addText( [{text:'Table Examples: ', options:ttsTitleText},{text:'autoPageLineWeight:0', options:ttsMultiOpts}], {x:0.5, y:0.13, w:3} );
	slide.addTable( arrText, { x:0.50, y:0.6, w:4, margin:5, border:'CFCFCF', autoPage:true } );

	slide.addText( [{text:'Table Examples: ', options:ttsTitleText},{text:'autoPageLineWeight:0.5', options:ttsMultiOpts}], {x:5.0, y:0.13, w:3} );
	slide.addTable( arrText, { x:4.75, y:0.6, w:4, margin:5, border:'CFCFCF', autoPage:true, autoPageLineWeight:0.5 } );

	slide.addText( [{text:'Table Examples: ', options:ttsTitleText},{text:'autoPageLineWeight:-0.5', options:ttsMultiOpts}], {x:9.0, y:0.13, w:3} );
	slide.addTable( arrText, { x:9.10, y:0.6, w:4, margin:5, border:'CFCFCF', autoPage:true, autoPageLineWeight:-0.5 } );

	pptx.writeFile({ fileName: 'PptxGenJS_TTSMulti_'+getTimestamp() });
}

function buildDataTable() {
	// STEP 1:
	$('#tabAutoPaging tbody').empty();

	// STEP 2:
	for (var idx=0; idx<$('#numTab2SlideRows').val(); idx++) {
		var strHtml = '<tr>'
			+ '<td style="text-align:center">' + (idx+1) + '</td>'
			+ '<td>' + gArrNamesL[ Math.floor(Math.random()*10) ] + '</td>'
			+ '<td>' + gArrNamesF[ Math.floor(Math.random()*10) ] + '</td>'
			+ '<td>Text:<br>' + gStrLoremIpsum.substring( 0, (Math.floor(Math.random()*10)+2)*130 ) + '</td>'
			+ '</tr>';
		$('#tabAutoPaging tbody').append( strHtml );
	}

	// STEP 3: Add some style to table for testing
	// TEST Padding
	$('#tabAutoPaging thead th').css('padding','10px 5px');
	// TEST font-size/auto-paging
	$('#tabAutoPaging tbody tr:first-child td:last-child').css('font-size','12px');
	$('#tabAutoPaging tbody tr:last-child td:last-child').css('font-size','16px');
}

function table2slidesDemoForTab(inTabId,inOpts) {
	var pptx = new PptxGenJS();
	pptx.tableToSlides(inTabId,(inOpts||null));
	pptx.writeFile({ fileName:  inTabId+'_'+getTimestamp() });
}

function table2slidesBullets() {
	var pptx = new PptxGenJS();
	pptx.tableToSlides('tableWithBullets');
	pptx.writeFile({ fileName:  'tabBullets_'+getTimestamp() });
}

/* DESC: Test for backward compatibility with Slide Masters defined in `pptxgen.masters.js` */
function testOnly_LegacyMasterSlides() {
	// TEST-ONLY: DO NOT USE/COPY ME!!
	var pptx = new PptxGenJS();
	pptx.layout = 'LAYOUT_WIDE';
	var slide = pptx.addSlide( pptx.masters.TITLE_SLIDE  );
	pptx.writeFile({ fileName: 'Demo-LegacyMasterSlides_'+getTimestamp() });
}
