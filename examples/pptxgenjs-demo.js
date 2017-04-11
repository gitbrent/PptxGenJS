/**
 * NAME: pptxgenjs-demo.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: Apr 10, 2017
 * DESC: Common test/demo slides for all library features
 * DEPS: Loaded by `pptxgenjs-demo.js` and `nodejs-demo.js`
**/

// Detect Node.js
var NODEJS = ( typeof module !== 'undefined' && module.exports );
// Constants
var CUST_NAME = 'S.T.A.R. Laboratories';
var USER_NAME = 'Barry Allen';
var COLOR_RED = 'FF0000';
var COLOR_AMB = 'F2AF00';
var COLOR_GRN = '7AB800';
var COLOR_CRT = 'AA0000';
var COLOR_BLU = '0088CC';
//
var ARRSTRBITES = [130];
var CHARSPERLINE = 130; // "Open Sans", 13px, 900px-colW = ~19 words/line ~130 chars/line
// FYI: 3086 chars
var gStrLorumIpsum = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin condimentum dignissim velit vel luctus. Donec feugiat ipsum quis tempus blandit. Donec mattis mauris vel est dictum interdum. Pellentesque imperdiet nibh vitae porta ornare. Fusce non nisl lacus. Curabitur ut mattis dui. Ut pulvinar urna velit, vitae aliquam neque pulvinar eu. Fusce eget tellus eu lorem finibus mattis. Nunc blandit consequat arcu. Ut sed pharetra tortor, nec finibus ipsum. Pellentesque a est vitae ligula imperdiet rhoncus. Ut quis hendrerit tellus. Phasellus non malesuada mi. Suspendisse ullamcorper tristique odio fermentum elementum. Phasellus mattis mollis mauris, non mattis ligula dapibus quis. Quisque pretium metus massa. '
+ 'Curabitur condimentum consequat felis, id rutrum velit cursus vel. Proin nulla est, posuere in velit at, faucibus dignissim diam. Quisque quis erat euismod, malesuada erat eu, congue nisi. Ut risus lectus, auctor at libero sit amet, accumsan ultricies est. Donec eget iaculis enim. Nunc ac egestas tellus, nec efficitur magna. Sed nec nisl ut augue laoreet sollicitudin vitae nec quam. Vestibulum pretium nisl bibendum, tempor velit eu, semper velit. Nulla facilisi. Aenean quis purus sagittis, dapibus nibh eget, ornare nunc. Donec posuere erat quis ipsum facilisis, quis porttitor dui cursus. Etiam convallis arcu sapien, vitae placerat diam molestie sit amet. Vivamus sapien augue, porta sed tortor ut, molestie ornare nisl. Nullam sed mi turpis. Donec sed finibus risus. '
+ 'Nunc interdum semper mauris quis vehicula. Phasellus in nisl faucibus, pellentesque massa vel, faucibus urna. Proin sed tortor lorem. Curabitur eu nisi semper, placerat tellus sed, varius nulla. Etiam luctus ac purus nec aliquet. Phasellus nisl metus, dictum ultricies justo a, laoreet consectetur risus. Vestibulum vulputate in felis ac blandit. Aliquam erat volutpat. Sed quis ultrices lectus. '
+ 'Curabitur at scelerisque elit, a bibendum nisi. Integer facilisis ex dolor, vel gravida metus vestibulum ac. Aliquam condimentum fermentum rhoncus. Nunc tortor arcu, condimentum non ex consequat, porttitor maximus est. Duis semper risus odio, quis feugiat sem elementum nec. Nam mattis nec dui sit amet volutpat. Sed facilisis, nunc quis porta consequat, ante mi tincidunt massa, eget euismod sapien nunc eget sem. Curabitur orci neque, eleifend at mattis quis, malesuada ac nibh. Vestibulum sed laoreet dolor, ac facilisis urna. Vestibulum luctus id nulla at auctor. Nunc pharetra massa orci, ut pharetra metus faucibus eget.'
+ 'Etiam eleifend, tellus id lobortis molestie, sem magna elementum dui, dapibus ullamcorper nisl enim ac urna. Nam posuere ullamcorper tellus, ac blandit nulla vestibulum nec. Vestibulum ornare, ligula quis aliquet cursus, metus nisi congue nulla, vitae posuere elit mauris at justo. Nullam ut fermentum arcu, nec laoreet ligula. Morbi quis consectetur nisl, nec consectetur justo. Curabitur eget eros hendrerit, ullamcorper dolor non, aliquam elit. Aliquam mollis justo vel aliquam interdum. Aenean bibendum rhoncus ante a commodo. Vestibulum bibendum sapien a accumsan pharetra... '
+ 'Curabitur condimentum consequat felis, id rutrum velit cursus vel. Proin nulla est, posuere in velit at, faucibus dignissim diam. Quisque quis erat euismod, malesuada erat eu, congue nisi. Ut risus lectus, auctor at libero sit amet, accumsan ultricies est. Donec eget iaculis enim. Nunc ac egestas tellus, nec efficitur magna. Sed nec nisl ut augue laoreet sollicitudin vitae nec quam. Vestibulum pretium nisl bibendum, tempor velit eu, semper velit. Nulla facilisi. Aenean quis purus sagittis, dapibus nibh eget, ornare nunc. Donec posuere erat quis ipsum facilisis, quis porttitor dui cursus. Etiam convallis arcu sapien, vitae placerat diam molestie sit amet. Vivamus sapien augue, porta sed tortor ut, molestie ornare nisl. Nullam sed mi turpis. Donec sed finibus risus. '
+ 'Nunc interdum semper mauris quis vehicula. Phasellus in nisl faucibus, pellentesque massa vel, faucibus urna. Proin sed tortor lorem. Curabitur eu nisi semper, placerat tellus sed, varius nulla. Etiam luctus ac purus nec aliquet. Phasellus nisl metus, dictum ultricies justo a, laoreet consectetur risus. Vestibulum vulputate in felis ac blandit. Aliquam erat volutpat. Sed quis ultrices lectus. '
+ 'Curabitur at scelerisque elit, a bibendum nisi. Integer facilisis ex dolor, vel gravida metus vestibulum ac. Aliquam condimentum fermentum rhoncus. Nunc tortor arcu, condimentum non ex consequat, porttitor maximus est. Duis semper risus odio, quis feugiat sem elementum nec. Nam mattis nec dui sit amet volutpat. Sed facilisis, nunc quis porta consequat, ante mi tincidunt massa, eget euismod sapien nunc eget sem. Curabitur orci neque, eleifend at mattis quis, malesuada ac nibh. Vestibulum sed laoreet dolor, ac facilisis urna. Vestibulum luctus id nulla at auctor. Nunc pharetra massa orci, ut pharetra metus faucibus eget.'
+ 'Etiam eleifend, tellus id lobortis molestie, sem magna elementum dui, dapibus ullamcorper nisl enim ac urna. Nam posuere ullamcorper tellus, ac blandit nulla vestibulum nec. Vestibulum ornare, ligula quis aliquet cursus, metus nisi congue nulla, vitae posuere elit mauris at justo. Nullam ut fermentum arcu, nec laoreet ligula. Morbi quis consectetur nisl, nec consectetur justo. Curabitur eget eros hendrerit, ullamcorper dolor non, aliquam elit. Aliquam mollis justo vel aliquam interdum. Aenean bibendum rhoncus ante a commodo. Vestibulum bibendum sapien a accumsan pharetra.';
var gArrNamesF = ['Markiplier','Jack','Brian','Paul','Ev','Ann','Michelle','Jenny','Lara','Kathryn'];
var gArrNamesL = ['Johnson','Septiceye','Lapston','Lewis','Clark','Griswold','Hart','Cube','Malloy','Capri'];
var gStrHello = 'BONJOUR - CIAO - GUTEN TAG - HELLO - HOLA - NAMASTE - OLÀ - ZDRAS-TVUY-TE - こんにちは - 你好';
var colors = ['FF0000','AB00CD','00FF00','00AA00','003300','330033','990099','33FFFF','AA33CC','336699'];
var fonts = ['Arial','Courier New','Times','Verdana'];
//
var optsTitle = { color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] };
var optsSubTitle = { x:0.5, y:0.7, cx:4, cy:0.3, font_size:18, font_face:'Arial', color:'0088CC', fill:'FFFFFF' };
var textTitle = { font_size:14, color:'0088CC', bold:true };
var textSubtt = { font_size:13, color:'9F9F9F' };

// ==================================================================================================================

function getTimestamp() {
	var dateNow = new Date();
	var dateMM = dateNow.getMonth() + 1; dateDD = dateNow.getDate(); dateYY = dateNow.getFullYear(), h = dateNow.getHours(); m = dateNow.getMinutes();
	return dateNow.getFullYear() +''+ (dateMM<=9 ? '0' + dateMM : dateMM) +''+ (dateDD<=9 ? '0' + dateDD : dateDD) + (h<=9 ? '0' + h : h) + (m<=9 ? '0' + m : m);
}

// ==================================================================================================================

function runEveryTest() {
	execGenSlidesFuncs( ['Table', 'Text', 'Image', 'Media', 'Shape', 'Master'] );
	if ( typeof table2slides1 !== 'undefined' ) table2slides1();
}

function execGenSlidesFuncs(type) {
	// STEP 1: Instantiate new PptxGenJS object
	if ( NODEJS ) {
		var pptx = require("pptxgenjs");
		//var pptx = require('../dist/pptxgen.js'); // for LOCAL TESTING
	}
	else {
		var pptx = new PptxGenJS();
	}

	pptx.setLayout('LAYOUT_WIDE');

	pptx.setAuthor('Brent Ely');
	pptx.setCompany('S.T.A.R. Laboratories');
	pptx.setRevision('15');
	pptx.setSubject('PptxGenJS Test Suite Export');
	pptx.setTitle('PptxGenJS Test Suite Presentation');

	// STEP 2: Run requested test
	var arrTypes = ( typeof type === 'string' ? [type] : type );
	arrTypes.forEach(function(type,idx){
		eval( 'genSlides_'+type+'(pptx)' );
	});

	// LAST: Export Presentation
	if ( !NODEJS ) pptx.save('Demo-'+type+'_'+getTimestamp());
}

// ==================================================================================================================

function genSlides_Table(pptx) {
	// SLIDE 1: Table text alignment and cell styles
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Table Examples 1', opts:optsTitle }] ], { x:0.5, y:0.13, w:12.5, h:0.3 } ); // `opts` = legacy test

		// DEMO: align/valign -------------------------------------------------------------------------
		var objOpts1 = { x:0.5, y:0.7, font_size:18, font_face:'Arial', color:'0088CC' };
		slide.addText('Cell Text Alignment:', objOpts1);

		var arrTabRows = [
			[
				{ text: 'Top Lft', options: { valign:'top',    align:'left'  , font_face:'Arial'   } },
				{ text: 'Top Ctr', options: { valign:'t'  ,    align:'center', font_face:'Courier' } },
				{ text: 'Top Rgt', options: { valign:'t'  ,    align:'right' , font_face:'Verdana' } }
			],
			[
				{ text: 'Ctr Lft', options: { valign:'middle', align:'left' } },
				{ text: 'Ctr Ctr', options: { valign:'center', align:'ctr'  } },
				{ text: 'Ctr Rgt', options: { valign:'c'     , align:'r'    } }
			],
			[
				{ text: 'Btm Lft', options: { valign:'bottom', align:'l' } },
				{ text: 'Btm Ctr', options: { valign:'btm',    align:'c' } },
				{ text: 'Btm Rgt', options: { valign:'b',      align:'r' } }
			]
		];
		slide.addTable(
			arrTabRows, { x:0.5, y:1.1, w:5.0 },
			{ rowH:0.75, fill:'F7F7F7', font_size:14, color:'363636', border:{pt:'1', color:'BBCCDD'} }
		);
		// Pass default cell style as tabOpts, then just style/override individual cells as needed

		// DEMO: cell styles --------------------------------------------------------------------------
		var objOpts2 = { x:6.0, y:0.7, font_size:18, font_face:'Arial', color:'0088CC' };
		slide.addText('Cell Styles:', objOpts2);

		var arrTabRows = [
			[
				{ text: 'White',  options: { fill:'6699CC', color:'FFFFFF' } },
				{ text: 'Yellow', options: { fill:'99AACC', color:'FFFFAA' } },
				{ text: 'Pink',   options: { fill:'AACCFF', color:'E140FE' } }
			],
			[
				{ text: '12pt', options: { fill:'FF0000', font_size:12 } },
				{ text: '20pt', options: { fill:'00FF00', font_size:20 } },
				{ text: '28pt', options: { fill:'0000FF', font_size:28 } }
			],
			[
				{ text: 'Bold',      options: { fill:'003366', bold:true } },
				{ text: 'Underline', options: { fill:'336699', underline:true } },
				{ text: '10pt Pad',  options: { fill:'6699CC', marginPt:10 } }
			]
		];
		slide.addTable(
			arrTabRows, { x:6.0, y:1.1, w:7.0 },
			{ rowH:0.75, fill:'F7F7F7', color:'FFFFFF', font_size:16, valign:'center', align:'ctr', border:{pt:'1', color:'FFFFFF'} }
		);

		// DEMO: Row/Col Width/Heights ----------------------------------------------------------------
		var objOpts3 = { x:0.5, y:3.6, font_size:18, font_face:'Arial', color:'0088CC' };
		slide.addText('Row/Col Heights/Widths:', objOpts3);

		var arrTabRows = [
			[ {text:'1x1'}, {text:'2x1'}, { text:'2.5x1' }, { text:'3x1' }, { text:'4x1' } ],
			[ {text:'1x2'}, {text:'2x2'}, { text:'2.5x2' }, { text:'3x2' }, { text:'4x2' } ]
		];
		slide.addTable( arrTabRows,
			{
				x:0.5, y:4.0,
				rowH: [1, 2], colW: [1, 2, 2.5, 3, 4],
				fill:'F7F7F7', color:'6c6c6c',
				font_size:14, valign:'center', align:'ctr',
				border:{pt:'1', color:'BBCCDD'}
			}
		);
	}

	// SLIDE 2: Table row/col-spans
	{
		var slide = pptx.addNewSlide();
		// 2: Slide title
		slide.addTable(
			[ [{ text:'Table Examples 2 [QA: this tables x,y,w,h all using %]', options:{ color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] } }] ],
			{ x:'5%', y:'2%', w:'90%', h:'4%' }
		);

		// DEMO: Rowspans/Colspans ----------------------------------------------------------------
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle));
		slide.addText('Colspans/Rowspans:', optsSub);

		var tabOpts1 = { x:0.5, y:1.1, w:'90%', h:2, fill:'F5F5F5', color:'3D3D3D', font_size:16, border:{pt:4, color:'FFFFFF'}, align:'c', valign:'m' };
		var arrTabRows1 = [
			[
				 { text:'A1\nA2', options:{rowspan:2, fill:'99FFCC'} }
				,{ text:'B1' }
				,{ text:'C1 -> D1', options:{colspan:2, fill:'99FFCC'} }
				,{ text:'E1' }
				,{ text:'F1\nF2\nF3', options:{rowspan:3, fill:'99FFCC'} }
			]
			,[       'B2', 'C2', 'D2', 'E2' ]
			,[ 'A3', 'B3', 'C3', 'D3', 'E3' ]
		];
		// NOTE: Follow HTML conventions for colspan/rowspan cells - cells spanned are left out of arrays - see above
		// The table above has 6 columns, but each of the 3 rows has 4-5 elements as colspan/rowspan replacing the missing ones
		// (e.g.: there are 5 elements in the first row, and 6 in the second)
		slide.addTable( arrTabRows1, tabOpts1 );

		var tabOpts2 = { x:0.5, y:3.3, w:12.4, h:1.5, font_size:14, font_face:'Courier', align:'center', valign:'middle', fill:'F9F9F9', border:{pt:'1',color:'c7c7c7'}};
		var arrTabRows2 = [
			[
				{ text:'A1\n--\nA2', options:{rowspan:2, fill:'99FFCC'} },
				{ text:'B1\n--\nB2', options:{rowspan:2, fill:'99FFCC'} },
				{ text:'C1 -> D1',   options:{colspan:2, fill:'9999FF'} },
				{ text:'E1 -> F1',   options:{colspan:2, fill:'9999FF'} },
				'G1'
			],
			[ 'C2','D2','E2','F2','G2' ]
		];
		slide.addTable( arrTabRows2, tabOpts2 );

		var tabOpts3 = {x:0.5, y:5.15, w:6.25, h:2, margin:0.25, align:'center', valign:'middle', font_size:16, border:{pt:'1',color:'c7c7c7'}, fill:'F1F1F1' }
		var arrTabRows3 = [
			[ {text:'A1\nA2\nA3', options:{rowspan:3, fill:'FFFCCC'}}, {text:'B1\nB2', options:{rowspan:2, fill:'FFFCCC'}}, 'C1' ],
			[ 'C2' ],
			[ { text:'B3 -> C3', options:{colspan:2, fill:'99FFCC'} } ]
		];
		slide.addTable(arrTabRows3, tabOpts3);

		var tabOpts4 = {x:7.4, y:5.15, w:5.5, h:2, margin:0, align:'center', valign:'middle', font_size:16, border:{pt:'1',color:'c7c7c7'}, fill:'F2F9FC' }
		var arrTabRows4 = [
			[ 'A1', {text:'B1\nB2', options:{rowspan:2, fill:'FFFCCC'}}, {text:'C1\nC2\nC3', options:{rowspan:3, fill:'FFFCCC'}} ],
			[ 'A2' ],
			[ { text:'A3 -> B3', options:{colspan:2, fill:'99FFCC'} } ]
		];
		slide.addTable(arrTabRows4, tabOpts4);
	}

	// SLIDE 3: Super rowspan/colspan demo
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Table Examples 3', opts:optsTitle }] ], { x:0.5, y:0.13, w:'94%', h:0.3 } ); // `opts` = legacy test

		// DEMO: Rowspans/Colspans ----------------------------------------------------------------
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle));
		slide.addText('Extreme Colspans/Rowspans:', optsSub);

		var optsRowspan2 = {rowspan:2, fill:'99FFCC'};
		var optsRowspan3 = {rowspan:3, fill:'99FFCC'};
		var optsRowspan4 = {rowspan:4, fill:'99FFCC'};
		var optsRowspan5 = {rowspan:5, fill:'99FFCC'};
		var optsColspan2 = {colspan:2, fill:'9999FF'};
		var optsColspan3 = {colspan:3, fill:'9999FF'};
		var optsColspan4 = {colspan:4, fill:'9999FF'};
		var optsColspan5 = {colspan:5, fill:'9999FF'};

		var arrTabRows5 = [
			[
				'A1','B1','C1','D1','E1','F1','G1','H1',
				{ text:'I1\n-\nI2\n-\nI3\n-\nI4\n-\nI5', options:optsRowspan5 },
				{ text:'J1 -> K1 -> L1 -> M1 -> N1', options:optsColspan5 }
			],
			[
				{ text:'A2\n--\nA3', options:optsRowspan2 },
				{ text:'B2 -> C2 -> D2',   options:optsColspan3 },
				{ text:'E2 -> F2',   options:optsColspan2 },
				{ text:'G2\n-\nG3\n-\nG4', options:optsRowspan3 },
				'H2',
				'J2','K2','L2','M2','N2'
			],
			[
				{ text:'B3\n-\nB4\n-\nB5', options:optsRowspan3 },
				'C3','D3','E3','F3', 'H3', 'J3','K3','L3','M3','N3'
			],
			[
				{ text:'A4\n--\nA5', options:optsRowspan2 },
				{ text:'C4 -> D4 -> E4 -> F4', options:optsColspan4 },
				'H4',
				{ text:'J4 -> K4 -> L4', options:optsColspan3 },
				{ text:'M4\n--\nM5', options:optsRowspan2 },
				{ text:'N4\n--\nN5', options:optsRowspan2 },
			],
			[
				'C5','D5','E5','F5',
				{ text:'G5 -> H5', options:{colspan:2, fill:'9999FF'} },
				'J5','K5','L5'
			]
		];

		var taboptions5 = { x:0.6, y:1.3, w:'90%', h:5.5, margin:0, font_size:14, align:'c', valign:'m', border:{pt:'1'} };

		slide.addTable(arrTabRows5, taboptions5);
	}

	// SLIDE 4: Cell Formatting / Cell Margins
	{
		var slide = pptx.addNewSlide();
		// 2: Slide title
		slide.addTable(
			[ [{ text:'Table Examples 4', options:{ color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] } }] ],
			{ x:0.5, y:0.13, w:12.5, h:0.3 }
		);

		// Cell Margins
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle));
		slide.addText('Cell Margins:', optsSub);

		slide.addTable( [['margin:0']],           { x:0.5, y:1.1, margin:0,           w:1.2, fill:'FFFCCC' } );
		slide.addTable( [['margin:[0,0,0,20]']],  { x:2.5, y:1.1, margin:[0,0,0,20],  w:2.0, fill:'FFFCCC', align:'r' } );
		slide.addTable( [['margin:5']],           { x:5.5, y:1.1, margin:5,           w:1.0, fill:'F1F1F1' } );
		slide.addTable( [['margin:[40,5,5,20]']], { x:7.5, y:1.1, margin:[40,5,5,20], w:2.2, fill:'F1F1F1' } );
		slide.addTable( [['margin:[80,5,5,10]']], { x:10.5,y:1.1, margin:[80,5,5,10], w:2.2, fill:'F1F1F1' } );

		// Complex/Compound border
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle)); optsSub.y = 2.6;
		slide.addText('Complex Cell Border:', optsSub);
		var arrBorder = [ {color:'FF0000',pt:1}, {color:'00ff00',pt:3}, {color:'0000ff',pt:5}, {color:'9e9e9e',pt:7} ];
		slide.addTable( [['Borders!']], { x:0.5, y:3.0, w:12.3, rowH:2.0, fill:'F5F5F5', color:'3D3D3D', font_size:18, border:arrBorder, align:'c', valign:'c' } );

		// Invalid char check
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle)); optsSub.y = 5.7;
		slide.addText('Escaped Invalid Chars:', optsSub);
		var arrTabRows3 = [['<', '>', '"', "'", '&', 'plain']];
		slide.addTable( arrTabRows3, { x:0.5, y:6.0, w:12.3, rowH:0.5, fill:'F5F5F5', color:'3D3D3D', border:'FFFFFF', align:'c', valign:'c' } );

	}

	// SLIDE 5: Cell Word-Level Formatting
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Table Examples 5', options:optsTitle }] ], { x:0.5, y:0.13, w:12.5, h:0.3 } );
		slide.addText(
			'The following textbox and table cell use the same array of text/options objects, making word-level formatting familiar and consistent across the library.',
			{ x:0.5, y:0.5, w:'95%', h:0.5, margin:0.1, font_size:14 }
		);
		slide.addText("[\n"
			+ "  { text:'1st line', options:{ font_size:24, color:'99ABCC', align:'r', breakLine:true } },\n"
			+ "  { text:'2nd line', options:{ font_size:36, color:'FFFF00', align:'c', breakLine:true } },\n"
			+ "  { text:'3rd line', options:{ font_size:48, color:'0088CC', align:'l' } }\n"
			+ "]",
			{ x:1, y:1.1, w:11, h:1.5, margin:0.1, font_face:'Courier', font_size:14, fill:'F1F1F1', color:'333333' }
		);

		// Textbox: Text word-level formatting
		slide.addText('Textbox:', { x:1, y:2.8, w:3, font_size:18, font_face:'Arial', color:'0088CC' });

		var arrTextObjects = [
			{ text:'1st line', options:{ font_size:24, color:'99ABCC', align:'r', breakLine:true } },
			{ text:'2nd line', options:{ font_size:36, color:'FFFF00', align:'c', breakLine:true } },
			{ text:'3rd line', options:{ font_size:48, color:'0088CC', align:'l' } }
		];
		slide.addText( arrTextObjects, { x:2.5, y:2.8, w:9, h:2, margin:0.1, fill:'232323' } );

		// Table cell: Use the exact same code from addText to do the same word-level formatting within a cell
		slide.addText('Table:', { x:1, y:5, w:3, font_size:18, font_face:'Arial', color:'0088CC' });

		var opts2 = { x:2.5, y:5, w:9, h:2, align:'center', valign:'middle', colW:[1.5,1.5,6], border:{pt:'1'}, fill:'F1F1F1' }
		var arrTabRows = [
			[
				{ text:'Cell 1A',       options:{font_face:'Arial'  } },
				{ text:'Cell 1B',       options:{font_face:'Courier'} },
				{ text: arrTextObjects, options:{fill:'232323'      } }
			]
		];
		slide.addTable(arrTabRows, opts2);
	}

	// SLIDE 6: Cell Word-Level Formatting
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [{ text:'Table Examples 6', options:{ color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] } }], { x:0.5, y:0.13, w:12.5, h:0.3 } );

		var optsSub = JSON.parse(JSON.stringify(optsSubTitle));
		slide.addText('Table Cell Word-Level Formatting:', optsSub);

		// EX 1:
		var arrCell1 = [{ text:'Cell 1A', options:{ color:'0088cc' } }];
		var arrCell2 = [{ text:'Red ', options:{color:'FF0000'} }, { text:'Green ', options:{color:'00FF00'} }, { text:'Blue', options:{color:'0000FF'} }];
		var arrCell3 = [{ text:'Bullets\nBullets\nBullets', options:{ color:'0088cc', bullet:true } }];
		var arrCell4 = [{ text:'Numbers\nNumbers\nNumbers', options:{ color:'0088cc', bullet:{type:'number'} } }];
		var arrTabRows = [
			[{ text:arrCell1 }, { text:arrCell2, options:{valign:'m'} }, { text:arrCell3, options:{valign:'m'} }, { text:arrCell4, options:{valign:'b'} }]
		];
		slide.addTable( arrTabRows, { x:0.6, y:1.25, w:12, h:3, font_size:24, border:{pt:'1'}, fill:'F1F1F1' } );

		// EX 2:
		slide.addTable(
			[
				{ text:[
						{ text:'I am a text object with bullets ', options:{color:'CC0000', bullet:{code:'2605'}} },
						{ text:'and i am the next text object'   , options:{color:'00CD00', bullet:{code:'25BA'}} },
						{ text:'Final text object w/ bullet:true', options:{color:'0000AB', bullet:true} }
				]},
				{ text:[
					{ text:'Cell', options:{font_size:36, align:'l', breakLine:true} },
					{ text:'#2',   options:{font_size:60, align:'r', color:'CD0101'} }
				]},
				{ text:[
					{ text:'Cell', options:{font_size:36, font_face:'Courier', color:'dd0000', breakLine:true} },
					{ text:'#'   , options:{font_size:60, color:'8648cd'} },
					{ text:'3'   , options:{font_size:60, color:'33ccef'} }
				]}
			],
			{ x:0.6, y:4.75, w:12, h:2, font_size:24, colW:[8,2,2], valign:'m', border:{pt:'1'}, fill:'F1F1F1' }
		);
	}

	// SLIDE 7+: Table auto-paging
	// ======== -----------------------------------------------------------------------------------
	{
		var arrRows = [];
		var arrText = [];
		for (var idx=0; idx<gArrNamesF.length; idx++) {
			var strText = ( idx == 0 ? gStrLorumIpsum.substring(0,100) : gStrLorumIpsum.substring(idx*100,idx*200) );
			arrRows.push( [idx, gArrNamesF[idx], strText] );
			arrText.push( [strText] );
		}

		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:textTitle},{text:'Auto-Paging Example', options:textSubtt}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:0.5, y:0.6, colW:[0.75,1.75,10], margin:2, border:'CFCFCF' } );

		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:textTitle},{text:'Smaller Table Area', options:textSubtt}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:3.0, y:0.6, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF' } );

		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:textTitle},{text:'Test: Correct starting Y location upon paging', options:textSubtt}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:3.0, y:4.0, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF' } );

		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:textTitle},{text:'Test: `{ newPageStartY: 1.5 }`', options:textSubtt}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:3.0, y:4.0, newPageStartY:1.5, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF' } );

		var slide = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, {bkgd:'CCFFCC'} );
		slide.addText( [{text:'Table Examples: ', options:textTitle},{text:'Master Page with Auto-Paging', options:textSubtt}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:1.0, y:0.6, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF' } );

		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:textTitle},{text:'Auto-Paging Disabled', options:textSubtt}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:1.0, y:0.6, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF', autoPage:false } );

		// lineWeight option demos
		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:textTitle},{text:'lineWeight:0', options:textSubtt}], {x:0.5, y:0.13, w:3} );
		slide.addTable( arrText, { x:0.50, y:0.6, w:4, margin:5, border:'CFCFCF', autoPage:true } );

		slide.addText( [{text:'Table Examples: ', options:textTitle},{text:'lineWeight:0.5', options:textSubtt}], {x:5.0, y:0.13, w:3} );
		slide.addTable( arrText, { x:4.75, y:0.6, w:4, margin:5, border:'CFCFCF', autoPage:true, lineWeight:0.5 } );

		slide.addText( [{text:'Table Examples: ', options:textTitle},{text:'lineWeight:-0.5', options:textSubtt}], {x:9.0, y:0.13, w:3} );
		slide.addTable( arrText, { x:9.10, y:0.6, w:4, margin:5, border:'CFCFCF', autoPage:true, lineWeight:-0.5 } );
	}
}

function genSlides_Media(pptx) {
	// SLIDE 1: Video and YouTube
	// ======== -----------------------------------------------------------------------------------
	var slide1 = pptx.addNewSlide();
	slide1.addTable( [ [{ text:'Media: Video Examples', opts:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

	slide1.addText('Video: m4v', { x:0.5, y:0.6, w:4.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:0.5, y:1.0, w:4.00, h:2.27, type:'video', path:'media/sample.m4v' });

	slide1.addText('Video: mpg', { x:5.5, y:0.6, w:3.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:5.5, y:1.0, w:3.00, h:2.05, type:'video', path:'media/sample.mpg' });

	slide1.addText('Video: mov', { x:9.4, y:0.6, w:3.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:9.4, y:1.0, w:3.00, h:1.71, type:'video', path:'media/sample.mov' });

	slide1.addText('Video: mp4', { x:0.5, y:3.6, w:4.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:0.5, y:4.0, w:4.00, h:3.00, type:'video', path:'media/sample.mp4'});

	slide1.addText('Video: avi', { x:5.5, y:3.6, w:3.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:5.5, y:4.0, w:3.00, h:2.25, type:'video', path:'media/sample.avi' });

	slide1.addText('Online: YouTube', { x:9.4, y:3.6, w:3.00, h:0.4, color:'0088CC' });
	// Provide the usual options (locations and size), then pass the embed code from YouTube (it's on every video page)
	slide1.addMedia({ x:9.4, y:4.0, w:3.00, h:2.25, type:'online', link:'https://www.youtube.com/embed/Dph6ynRVyUc' });

	// SLIDE 2: Audio / Pre-Encoded Video
	// ======== -----------------------------------------------------------------------------------
	var slide2 = pptx.addNewSlide();
	slide2.addTable( [ [{ text:'Media: Audio and Pre-Encoded Audio/Video Examples', opts:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

	slide2.addText('Audio: mp3', { x:0.5, y:0.6, w:4.00, h:0.4, color:'0088CC' });
	slide2.addMedia({ x:0.5, y:1.0, w:4.00, h:0.3, type:'audio', path:'media/sample.mp3' });

	slide2.addText('Audio: wav', { x:0.5, y:2.6, w:4.00, h:0.4, color:'0088CC' });
	slide2.addMedia({ x:0.5, y:3.0, w:4.00, h:0.3, type:'audio', path:'media/sample.wav' });

	//slide2.addText('Audio: Pre-Encoded mp3', { x:5.5, y:0.6, w:4.00, h:0.4, color:'0088CC' });
	//slide2.addMedia({ x:5.5, y:1.0, w:4.00, h:0.3, type:'audio', data:AUDIO_MP3 }); // Keynote=pass,O365=fail

	//slide2.addText('Video: Pre-Encoded mp4', { x:5.5, y:2.6, w:4.00, h:0.4, color:'0088CC' });
	//slide2.addMedia({ x:5.5, y:3.0, w:4.00, h:3.0, type:'video', data:VIDEO_MP4 }); // Keynote=pass,O365=fail
}

function genSlides_Image(pptx) {
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Image Examples: Misc Image Types', options:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

	slide.addText('Type: GIF', { x:0.5, y:0.6, w:2.5, h:0.4, color:'0088CC' });
	slide.addImage({ path:'images/cc_copyremix.gif', x:0.5, y:1.0, w:1.2, h:1.2 });

	slide.addText('Type: JPG', { x:0.5, y:3.0, w:2.5, h:0.4, color:'0088CC' });
	slide.addImage({ path:'images/cc_logo.jpg', x:0.5, y:3.5, w:5.0, h:3.7 });

	slide.addText('Type: PNG', { x:6.6, y:3.0, w:2.5, h:0.4, color:'0088CC' });
	slide.addImage({ path:'images/cc_license_comp.png', x:6.6, y:3.5, w:6.3, h:3.7 });

	slide.addText('Type: Anim-GIF', { x:3.5, y:0.6, w:2.5, h:0.4, color:'0088CC' });
	if (NODEJS) slide.addImage({ x:3.5, y:0.8, w:1.78, h:1.78, path:'images/anim_campfire.gif' });
	else        slide.addImage({ x:3.5, y:0.8, w:1.78, h:1.78, data:GIF_ANIM_FIRE });

	// Images can be pre-encoded into base64, so they do not have to be on the webserver etc. (saves generation time and resources!)
	// Also has the benefit of being able to be any type (path:images can only be exported as PNG)
	// NOTE: The 'data:' part of the encoded string is optional:

	slide.addText('Pre-Encoded PNG', { x:6.6, y:0.6, w:3.0, h:0.4, color:'0088CC' });
	slide.addImage({
		x:6.6, y:1.2, w:0.6, h:0.6,
		data:'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAAjcAAAI3AGf6F88AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAANVQTFRF////JLaSIJ+AIKqKKa2FKLCIJq+IJa6HJa6JJa6IJa6IJa2IJa6IJa6IJa6IJa6IJa6IJa6IJq6IKK+JKK+KKrCLLrGNL7KOMrOPNrSRN7WSPLeVQrmYRLmZSrycTr2eUb6gUb+gWsKlY8Wqbsmwb8mwdcy0d8y1e863g9G7hdK8htK9i9TAjNTAjtXBktfEntvKoNzLquDRruHTtePWt+TYv+fcx+rhyOvh0e7m1e/o2fHq4PTu5PXx5vbx7Pj18fr49fv59/z7+Pz7+f38/P79/f7+dNHCUgAAABF0Uk5TAAcIGBktSYSXmMHI2uPy8/XVqDFbAAABB0lEQVQ4y42T13qDMAyFZUKMbebp3mmbrnTvlY60TXn/R+oFGAyYzz1Xx/wylmWJqBLjUkVpGinJGXXliwSVEuG3sBdkaCgLPJMPQnQUDmo+jGFRPKz2WzkQl//wQvQoLPII0KuAiMjP+gMyn4iEFU1eAQCCiCU2fpCfFBVjxG18f35VOk7Swndmt9pKUl2++fG4qL2iqMPXpi8r1SKitDDne/rT8vPbRh2d6oC7n6PCLNx/bsEM0Edc5DdLAHD9tWueF9VJjmdP68DZ77iRkDKuuT19Hx3mx82MpVmo1Yfv+WXrSrxZ6slpiyes77FKif88t7Nh3C3nbFp327sHxz167uHtH/8/eds7gGsUQbkAAAAASUVORK5CYII='
	});

	// TEST: Ensure framework corrects for missing type header
	slide.addText('Pre-Encoded PNG', { x:9.6, y:0.6, w:3.0, h:0.4, color:'0088CC' });
	slide.addImage({
		x:9.8, y:1.2, w:0.8, h:0.8,
		data:'base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAAjcAAAI3AGf6F88AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAANVQTFRF////JLaSIJ+AIKqKKa2FKLCIJq+IJa6HJa6JJa6IJa6IJa2IJa6IJa6IJa6IJa6IJa6IJa6IJq6IKK+JKK+KKrCLLrGNL7KOMrOPNrSRN7WSPLeVQrmYRLmZSrycTr2eUb6gUb+gWsKlY8Wqbsmwb8mwdcy0d8y1e863g9G7hdK8htK9i9TAjNTAjtXBktfEntvKoNzLquDRruHTtePWt+TYv+fcx+rhyOvh0e7m1e/o2fHq4PTu5PXx5vbx7Pj18fr49fv59/z7+Pz7+f38/P79/f7+dNHCUgAAABF0Uk5TAAcIGBktSYSXmMHI2uPy8/XVqDFbAAABB0lEQVQ4y42T13qDMAyFZUKMbebp3mmbrnTvlY60TXn/R+oFGAyYzz1Xx/wylmWJqBLjUkVpGinJGXXliwSVEuG3sBdkaCgLPJMPQnQUDmo+jGFRPKz2WzkQl//wQvQoLPII0KuAiMjP+gMyn4iEFU1eAQCCiCU2fpCfFBVjxG18f35VOk7Swndmt9pKUl2++fG4qL2iqMPXpi8r1SKitDDne/rT8vPbRh2d6oC7n6PCLNx/bsEM0Edc5DdLAHD9tWueF9VJjmdP68DZ77iRkDKuuT19Hx3mx82MpVmo1Yfv+WXrSrxZ6slpiyes77FKif88t7Nh3C3nbFp327sHxz167uHtH/8/eds7gGsUQbkAAAAASUVORK5CYII='
	});

	// TEST: Ensure framework corrects for missing all header (Please DO NOT pass base64 data without the header! This is a junky test)
	//slide.addImage({ x:5.2, y:2.6, w:0.8, h:0.8, data:'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAAjcAAAI3AGf6F88AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAANVQTFRF////JLaSIJ+AIKqKKa2FKLCIJq+IJa6HJa6JJa6IJa6IJa2IJa6IJa6IJa6IJa6IJa6IJa6IJq6IKK+JKK+KKrCLLrGNL7KOMrOPNrSRN7WSPLeVQrmYRLmZSrycTr2eUb6gUb+gWsKlY8Wqbsmwb8mwdcy0d8y1e863g9G7hdK8htK9i9TAjNTAjtXBktfEntvKoNzLquDRruHTtePWt+TYv+fcx+rhyOvh0e7m1e/o2fHq4PTu5PXx5vbx7Pj18fr49fv59/z7+Pz7+f38/P79/f7+dNHCUgAAABF0Uk5TAAcIGBktSYSXmMHI2uPy8/XVqDFbAAABB0lEQVQ4y42T13qDMAyFZUKMbebp3mmbrnTvlY60TXn/R+oFGAyYzz1Xx/wylmWJqBLjUkVpGinJGXXliwSVEuG3sBdkaCgLPJMPQnQUDmo+jGFRPKz2WzkQl//wQvQoLPII0KuAiMjP+gMyn4iEFU1eAQCCiCU2fpCfFBVjxG18f35VOk7Swndmt9pKUl2++fG4qL2iqMPXpi8r1SKitDDne/rT8vPbRh2d6oC7n6PCLNx/bsEM0Edc5DdLAHD9tWueF9VJjmdP68DZ77iRkDKuuT19Hx3mx82MpVmo1Yfv+WXrSrxZ6slpiyes77FKif88t7Nh3C3nbFp327sHxz167uHtH/8/eds7gGsUQbkAAAAASUVORK5CYII=' });
	// NEGATIVE-TEST:
	//slide.addImage({ data:'images/doh_this_isnt_base64_data.gif',  x:0.5, y:0.5, w:1.0, h:1.0 });
}

function genSlides_Shape(pptx) {
	// SLIDE 1: Misc Shape Types (no text)
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Shape Examples 1: Misc Shape Types (no text)', options:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

	//slide.addShape(pptx.shapes.RECTANGLE,         { x:0.5, y:0.8, w:12.5,h:0.5, fill:'F9F9F9' });
	slide.addShape(pptx.shapes.RECTANGLE,         { x:0.5, y:0.8, w:1.5, h:3.0, fill:'FF0000' });
	slide.addShape(pptx.shapes.RECTANGLE,         { x:3.0, y:0.7, w:1.5, h:3.0, fill:'F38E00', rotate:45 });
	slide.addShape(pptx.shapes.OVAL,              { x:5.4, y:0.8, w:3.0, h:1.5, fill:{ type:'solid', color:'0088CC', alpha:25 } });
	slide.addShape(pptx.shapes.OVAL,              { x:7.7, y:1.4, w:3.0, h:1.5, fill:{ type:'solid', color:'FF00CC', alpha:50 }, rotate:90 });
	slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:10 , y:2.5, w:3.0, h:1.5, fill:'00FF00' });
	//
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:4.4, w:5.0, h:0.0, line:'FF0000', line_size:1 });
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:4.8, w:5.0, h:0.0, line:'FF0000', line_size:2, line_head:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:5.2, w:5.0, h:0.0, line:'FF0000', line_size:3, line_tail:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:5.6, w:5.0, h:0.0, line:'FF0000', line_size:4, line_head:'triangle', line_tail:'triangle' });
	//
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:0.4, y:4.3, w:6.0, h:3.0, fill:'0088CC', line:'000000', line_size:3 });
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:7.0, y:4.3, w:6.0, h:3.0, fill:'0088CC', line:'000000', flipH:true });

	// SLIDE 2: Misc Shape Types with Text
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Shape Examples 2: Misc Shape Types (with text)', options:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

	slide.addText('RECTANGLE',                  { shape:pptx.shapes.RECTANGLE,         x:0.5, y:0.8, w:1.5, h:3.0, fill:'FF0000', align:'c', font_size:14 });
	slide.addText('RECTANGLE (rotate:45)',      { shape:pptx.shapes.RECTANGLE,         x:3.0, y:0.7, w:1.5, h:3.0, fill:'F38E00', rotate:45, align:'c', font_size:14 });
	slide.addText('OVAL (alpha:25)',            { shape:pptx.shapes.OVAL,              x:5.4, y:0.8, w:3.0, h:1.5, fill:{ type:'solid', color:'0088CC', alpha:25 }, align:'c', font_size:14 });
	slide.addText('OVAL (rotate:90, alpha:50)', { shape:pptx.shapes.OVAL,              x:7.7, y:1.4, w:3.0, h:1.5, fill:{ type:'solid', color:'FF00CC', alpha:50 }, rotate:90, align:'c', font_size:14 });
	slide.addText('ROUNDED-RECTANGLE',          { shape:pptx.shapes.ROUNDED_RECTANGLE, x:10 , y:2.5, w:3.0, h:1.5, fill:'00FF00', align:'c', font_size:14 });
	//
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'c', x:4.15, y:4.40, w:5, h:0, line:'FF0000', line_size:1 });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'l', x:4.15, y:4.80, w:5, h:0, line:'FF0000', line_size:2, line_head:'triangle' });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'r', x:4.15, y:5.20, w:5, h:0, line:'FF0000', line_size:3, line_tail:'triangle' });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'c', x:4.15, y:5.60, w:5, h:0, line:'FF0000', line_size:4, line_head:'triangle', line_tail:'triangle' });
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE,    align:'c', x:0.40, y:4.30, w:6, h:3, fill:'0088CC', line:'000000', line_size:3 });
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE,    align:'c', x:7.00, y:4.30, w:6, h:3, fill:'0088CC', line:'000000', flipH:true });
}

function genSlides_Text(pptx) {
	// SLIDE 1: Line Break / Bullets
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Text Examples 1', options:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

		// LEFT COLUMN ------------------------------------------------------------

		// 1: Multi-Line Formatting
		slide.addText("Word-Level Formatting:", { x:0.5, y:0.5, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text:'1st\nline', options:{ font_size:24, font_face:'Courier New', color:'99ABCC', align:'r', breakLine:true } },
				{ text:'2nd line', options:{ font_size:36, font_face:'Arial',       color:'FFFF00', align:'c', breakLine:true } },
				{ text:'3rd line', options:{ font_size:48, font_face:'Verdana',     color:'0088CC', align:'l' } }
			],
			{ x:0.5, y:0.85, w:6, h:2.25, margin:0.1, fill:'232323' }
		);

		// 2: Line-Break Test
		slide.addText("Line-Breaks:", { x:0.5, y:3.35, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			'***Line-Break/Multi-Line Test***\n\nFirst line\nSecond line\nThird line',
			{ x:0.5, y:3.75, w:6, h:1.75, valign:'middle', align:'ctr', color:'6c6c6c', font_size:16, fill:'F2F2F2' }
		);

		// 3: Hyperlinks
		slide.addText("Hyperlinks:", { x:0.5, y:5.9, w:1.75, h:0.35, color:'0088CC' });
		slide.addText(
			[
				{ text:'Visit the ' },
				{ text:'PptxGenJS Project', options:{ hyperlink:{ url:'https://github.com/gitbrent/pptxgenjs', tooltip:'Visit Homepage' } } },
				{ text:' or ' },
				{ text:'(no tooltip)', options:{hyperlink:{url:'https://github.com/gitbrent'}} }
			],
			{ x:2.25, y:5.9, w:4.25, h:0.55, margin:0.1, fill:'F1F1F1', font_size:14 }
		);

		// 4: Text Effects: Shadow
		var shadowOpts = { type:'outer', color:'696969', blur:3, offset:10, angle:45, opacity:0.8 };
		slide.addText("Text Shadow:", { x:0.5, y:6.74, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			'Outer Shadow (blur:3, offset:10, angle:45, opacity:80%)',
			{ x:2.1, y:6.65, w:12, h:0.6, font_size:32, color:'0088cc', shadow:shadowOpts }
		);

		// RIGHT COLUMN ------------------------------------------------------------

		// 4: Regular bullets
		slide.addText("Bullets:", { x:7.5, y:0.65, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(12345                  , { x:8.0, y:1.05, w:'30%', h:0.5, color:'0000DE', font_face:"Courier New", bullet:true });
		slide.addText('String (number above)', { x:8.0, y:1.35, w:'30%', h:0.5, color:'00AA00', bullet:true });

		// 5: Bullets: Text With Line-Breaks
		slide.addText("Bullets with line-breaks:", { x:7.5, y:2.0, w:'40%', h:0.38, color:'0088CC' });
		slide.addText('Line 1\nLine 2\nLine 3', { x:8.0, y:2.4, w:'30%', h:1, color:'393939', font_size:16, fill:'F2F2F2', bullet:{type:'number'} });

		// 6: Bullets: With group of {text}
		slide.addText("Bullet with {text} objects:", { x:7.5, y:3.6, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text: 'big red words... ', options:{font_size:24, color:'FF0000'} },
				{ text: 'some green words.', options:{font_size:16, color:'00FF00'} }
			],
			{ x:8.0, y:4.0, w:5, h:0.5, margin:0.1, font_face:'Arial', bullet:{code:'25BA'} }
		);

		// 7: Bullets: Within a {text} object
		slide.addText("Bullet within {text} objects:", { x:7.5, y:4.8, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text:'I am a text object with bullets..', options:{bullet:{code:'2605'}, color:'CC0000'} },
				{ text:'and I am the next text object.'   , options:{bullet:{code:'25BA'}, color:'00CD00'} },
				{ text:'Default bullet text.. '           , options:{bullet:true, color:'696969'} },
				{ text:'Final text object w/ bullet:true.', options:{bullet:true, color:'0000AB'} }
			],
			{ x:8.0, y:5.15, w:'35%', h:1.4, color:'ABABAB', margin:1 }
		);
	}

	// SLIDE 2: Misc mess
	{
		var slide = pptx.addNewSlide();
		// Slide colors: bkgd/fore
		slide.back = '030303';
		slide.color = '9F9F9F';
		// Title
		slide.addTable( [ [{ text:'Text Examples 2', options:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		// Actual Textbox shape (can have any Height, can wrap text, etc.)
		slide.addText( 'Textbox (ctr/ctr)', { x:0.5, y:0.75, w:8.5, h:2.5, color:'FFFFFF', fill:'0000FF', valign:'c', align:'c', isTextBox:true } );
		slide.addText( 'Textbox (top/lft)', { x:10,  y:0.75, w:3.0, h:1.0, color:'FFFFFF', fill:'00CC00', valign:'t', align:'l', isTextBox:true } );
		slide.addText( 'Textbox (btm/rgt)', { x:10,  y:2.25, w:3.0, h:1.0, color:'FFFFFF', fill:'FF0000', valign:'b', align:'r', isTextBox:true } );

		slide.addText('Plain x/y coords', { x:10, y:3.5 });

		slide.addText('Escaped chars: \' " & < >', { x:10, y:4.5 });

		slide.addText('^ (50%/50%)', {x:'50%', y:'50%', w:2});

		// TEST: using {option}: Add text box with multiline options:
		slide.addText(
			[
				{ text:'word-level\nformatting', options:{ font_size:36, font_face:'Courier New', color:'99ABCC', align:'r', breakLine:true } },
				{ text:'...in the same textbox', options:{ font_size:48, font_face:'Arial', color:'FFFF00', align:'c' } }
			],
			{ x:0.5, y:4.1, w:8.5, h:2.0, margin:0.1, fill:'232323' }
		);

		var objOptions = {
			x:0, y:6.25, w:'100%', h:0.5, align:'c',
			font_face:'Arial', font_size:24, color:'00EC23', bold:true, italic:true, underline:true, margin:0, isTextBox:true
		};
		slide.addText('Arial 32pt, green, bold, italic, underline, margin:0, ctr', objOptions);

		slide.addText('Footer Bar: PptxGenJS version ' + pptx.getVersion() + ' (width:100%, valign:ctr)',
			{ x:0, y:6.75, w:'100%', h:0.75, fill:'f7f7f7', color:'666666', align:'center', valign:'middle' }
		);
	}
}

function genSlides_Master(pptx) {
	var slide1 = pptx.addNewSlide( pptx.masters.TITLE_SLIDE  );
	var slide2 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE );
	var slide3 = pptx.addNewSlide( pptx.masters.THANKS_SLIDE );

	var slide4 = pptx.addNewSlide( pptx.masters.TITLE_SLIDE,  { bkgd:'0088CC'} );
	var slide5 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:{ src:'images/title_bkgd_alt.jpg' } } );
	var slide6 = pptx.addNewSlide( pptx.masters.THANKS_SLIDE, { bkgd:'ffab33'} );
}

// ==================================================================================================================

if ( typeof module !== 'undefined' && module.exports ) {
	module.exports = runEveryTest;
}
