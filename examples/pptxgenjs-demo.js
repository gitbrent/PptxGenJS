/**
* NAME: pptxgenjs-demo.js
* AUTH: Brent Ely (https://github.com/gitbrent/)
* DATE: Aug 15, 2017
* DESC: Common test/demo slides for all library features
* DEPS: Loaded by `pptxgenjs-demo.js` and `nodejs-demo.js`
*/

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
var COLOR_UNK = 'A9A9A9';
//
var ARRSTRBITES = [130];
var CHARSPERLINE = 130; // "Open Sans", 13px, 900px-colW = ~19 words/line ~130 chars/line
// FYI: 3086 chars
var gStrLorumIpsum = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin condimentum dignissim velit vel luctus. Donec feugiat ipsum quis tempus blandit. Donec mattis mauris vel est dictum interdum. Pellentesque imperdiet nibh vitae porta ornare. Fusce non nisl lacus. Curabitur ut mattis dui. Ut pulvinar urna velit, vitae aliquam neque pulvinar eu. Fusce eget tellus eu lorem finibus mattis. Nunc blandit consequat arcu. Ut sed pharetra tortor, nec finibus ipsum. \
Pellentesque a est vitae ligula imperdiet rhoncus. Ut quis hendrerit tellus. Phasellus non malesuada mi. Suspendisse ullamcorper tristique odio fermentum elementum. Phasellus mattis mollis mauris, non mattis ligula dapibus quis. Quisque pretium metus massa. \
Curabitur condimentum consequat felis, id rutrum velit cursus vel. Proin nulla est, posuere in velit at, faucibus dignissim diam. Quisque quis erat euismod, malesuada erat eu, congue nisi. Ut risus lectus, auctor at libero sit amet, accumsan ultricies est. Donec eget iaculis enim. Nunc ac egestas tellus, nec efficitur magna. \
Sed nec nisl ut augue laoreet sollicitudin vitae nec quam. Vestibulum pretium nisl bibendum, tempor velit eu, semper velit. Nulla facilisi. Aenean quis purus sagittis, dapibus nibh eget, ornare nunc. Donec posuere erat quis ipsum facilisis, quis porttitor dui cursus. Etiam convallis arcu sapien, vitae placerat diam molestie sit amet. Vivamus sapien augue, porta sed tortor ut, molestie ornare nisl. Nullam sed mi turpis. Donec sed finibus risus. \
Nunc interdum semper mauris quis vehicula. Phasellus in nisl faucibus, pellentesque massa vel, faucibus urna. Proin sed tortor lorem. Curabitur eu nisi semper, placerat tellus sed, varius nulla. Etiam luctus ac purus nec aliquet. Phasellus nisl metus, dictum ultricies justo a, laoreet consectetur risus. Vestibulum vulputate in felis ac blandit. Aliquam erat volutpat. Sed quis ultrices lectus. \
Curabitur at scelerisque elit, a bibendum nisi. Integer facilisis ex dolor, vel gravida metus vestibulum ac. Aliquam condimentum fermentum rhoncus. Nunc tortor arcu, condimentum non ex consequat, porttitor maximus est. Duis semper risus odio, quis feugiat sem elementum nec. Nam mattis nec dui sit amet volutpat. Sed facilisis, nunc quis porta consequat, ante mi tincidunt massa, eget euismod sapien nunc eget sem. Curabitur orci neque, eleifend at mattis quis, \
malesuada ac nibh. Vestibulum sed laoreet dolor, ac facilisis urna. Vestibulum luctus id nulla at auctor. Nunc pharetra massa orci, ut pharetra metus faucibus eget.\
Etiam eleifend, tellus id lobortis molestie, sem magna elementum dui, dapibus ullamcorper nisl enim ac urna. Nam posuere ullamcorper tellus, ac blandit nulla vestibulum nec. Vestibulum ornare, ligula quis aliquet cursus, metus nisi congue nulla, vitae posuere elit mauris at justo. Nullam ut fermentum arcu, nec laoreet ligula. Morbi quis consectetur nisl, nec consectetur justo. Curabitur eget eros hendrerit, ullamcorper dolor non, aliquam elit. Aliquam mollis justo vel aliquam interdum.\ Aenean bibendum rhoncus ante a commodo. Vestibulum bibendum sapien a accumsan pharetra... \
Curabitur condimentum consequat felis, id rutrum velit cursus vel. Proin nulla est, posuere in velit at, faucibus dignissim diam. Quisque quis erat euismod, malesuada erat eu, congue nisi. Ut risus lectus, auctor at libero sit amet, accumsan ultricies est. Donec eget iaculis enim. Nunc ac egestas tellus, nec efficitur magna. Sed nec nisl ut augue laoreet sollicitudin vitae nec quam. Vestibulum pretium nisl bibendum, tempor velit eu, semper velit. Nulla facilisi. Aenean quis purus sagittis,\
dapibus nibh eget, ornare nunc. Donec posuere erat quis ipsum facilisis, quis porttitor dui cursus. Etiam convallis arcu sapien, vitae placerat diam molestie sit amet. Vivamus sapien augue, porta sed tortor ut, molestie ornare nisl. Nullam sed mi turpis. Donec sed finibus risus. \
Nunc interdum semper mauris quis vehicula. Phasellus in nisl faucibus, pellentesque massa vel, faucibus urna. Proin sed tortor lorem. Curabitur eu nisi semper, placerat tellus sed, varius nulla. Etiam luctus ac purus nec aliquet. Phasellus nisl metus, dictum ultricies justo a, laoreet consectetur risus. Vestibulum vulputate in felis ac blandit. Aliquam erat volutpat. Sed quis ultrices lectus. \
Curabitur at scelerisque elit, a bibendum nisi. Integer facilisis ex dolor, vel gravida metus vestibulum ac. Aliquam condimentum fermentum rhoncus. Nunc tortor arcu, condimentum non ex consequat, porttitor maximus est. Duis semper risus odio, quis feugiat sem elementum nec. Nam mattis nec dui sit amet volutpat. Sed facilisis, nunc quis porta consequat, ante mi tincidunt massa, eget euismod sapien nunc eget sem. Curabitur orci neque, eleifend at mattis quis, malesuada ac nibh. Vestibulum \
sed laoreet dolor, ac facilisis urna. Vestibulum luctus id nulla at auctor. Nunc pharetra massa orci, ut pharetra metus faucibus eget.\
Etiam eleifend, tellus id lobortis molestie, sem magna elementum dui, dapibus ullamcorper nisl enim ac urna. Nam posuere ullamcorper tellus, ac blandit nulla vestibulum nec. Vestibulum ornare, ligula quis aliquet cursus, metus nisi congue nulla, vitae posuere elit mauris at justo. Nullam ut fermentum arcu, nec laoreet ligula. Morbi quis consectetur nisl, nec consectetur justo. Curabitur eget eros hendrerit, ullamcorper dolor non, aliquam elit. Aliquam mollis justo vel aliquam interdum.\
Aenean bibendum rhoncus ante a commodo. Vestibulum bibendum sapien a accumsan pharetra.';
var gArrNamesF = ['Markiplier','Jack','Brian','Paul','Ev','Ann','Michelle','Jenny','Lara','Kathryn'];
var gArrNamesL = ['Johnson','Septiceye','Lapston','Lewis','Clark','Griswold','Hart','Cube','Malloy','Capri'];
var gStrHello = 'BONJOUR - CIAO - GUTEN TAG - HELLO - HOLA - NAMASTE - OLÀ - ZDRAS-TVUY-TE - こんにちは - 你好';
//var colors = ['FF0000','AB00CD','00FF00','00AA00','003300','330033','990099','33FFFF','AA33CC','336699'];
//var fonts = ['Arial','Courier New','Times','Verdana'];
//
var gOptsTitle = { color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] };
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
	execGenSlidesFuncs( ['Text', 'Chart', 'Table', 'Image', 'Media', 'Shape', 'Master'] );
	if ( typeof table2slides1 !== 'undefined' ) table2slides1();
}

function execGenSlidesFuncs(type) {
	// STEP 1: Instantiate new PptxGenJS object
	var pptx;
	if ( NODEJS ) {
		var fs = require('fs');
		if (fs.existsSync('../dist/pptxgen.js')) {
			pptx = require('../dist/pptxgen.js'); // for LOCAL TESTING
		}
		else {
			pptx = require("pptxgenjs");
		}
	}
	else {
		pptx = new PptxGenJS();
	}

	pptx.setLayout('LAYOUT_WIDE');

	pptx.setAuthor('Brent Ely');
	pptx.setCompany(CUST_NAME);
	pptx.setRevision('15');
	pptx.setSubject('PptxGenJS Test Suite Export');
	pptx.setTitle('PptxGenJS Test Suite Presentation');

	// STEP 2: Run requested test
	var arrTypes = ( typeof type === 'string' ? [type] : type );
	arrTypes.forEach(function(type,idx){ eval( 'genSlides_'+type+'(pptx)' ); });

	// LAST: Export Presentation
	if ( !NODEJS ) pptx.save('Demo-'+type+'_'+getTimestamp());
}

// ==================================================================================================================

function genSlides_Table(pptx) {
	// SLIDE 1: Table text alignment and cell styles
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Table Examples 1', opts:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5, h:0.3 } ); // `opts` = legacy test

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
		slide.addTable( [ [{ text:'Table Examples 3', opts:gOptsTitle }] ], { x:0.5, y:0.13, w:'94%', h:0.3 } ); // `opts` = legacy test

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

		slide.addTable( [{text:'number zero:', options:{margin:5}}, {text:0, options:{marginPt:5}}], { x:0.5, y:1.9, w:3, fill:'f2f9fc', colW:[2,1] } );
		slide.addTable( [{text:'text-obj margin:0', options:{margin:0}}], { x:4.0, y:1.9, w:2, fill:'f2f9fc' } );

		// Test margin option when using both plain and text object cells
		var arrTextObjects = [
			['Plain text','Cell 2',3],
			[
				{ text:'Text Objects', options:{ color:'99ABCC', align:'r' } },
				{ text:'2nd cell', options:{ color:'0000EE', align:'c' } },
				{ text:3, options:{ color:'0088CC', align:'l' } }
			]
		];
		slide.addTable( arrTextObjects, { x:0.5, y:2.7, w:12.25, margin:7, fill:'F1F1F1', border:{pt:1,color:'696969'} } );

		// Complex/Compound border
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle)); optsSub.y = 3.9;
		slide.addText('Complex Cell Border:', optsSub);
		var arrBorder = [ {color:'FF0000',pt:1}, {color:'00ff00',pt:3}, {color:'0000ff',pt:5}, {color:'9e9e9e',pt:7} ];
		slide.addTable( [['Borders!']], { x:0.5, y:4.3, w:12.3, rowH:1.5, fill:'F5F5F5', color:'3D3D3D', font_size:18, border:arrBorder, align:'c', valign:'c' } );

		// Invalid char check
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle)); optsSub.y = 6.1;
		slide.addText('Escaped Invalid Chars:', optsSub);
		var arrTabRows3 = [['<', '>', '"', "'", '&', 'plain']];
		slide.addTable( arrTabRows3, { x:0.5, y:6.5, w:12.3, rowH:0.5, fill:'F5F5F5', color:'3D3D3D', border:'FFFFFF', align:'c', valign:'c' } );

	}

	// SLIDE 5: Cell Word-Level Formatting
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Table Examples 5', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5, h:0.3 } );
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

function genSlides_Chart(pptx) {
	var LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
	var MONS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
	var QTRS = ['Q1','Q2','Q3','Q4'];

	var dataChartPieStat = [
		{
			name  : 'Project Status',
			labels: ['Red', 'Amber', 'Green', 'Unknown'],
			values: [8, 20, 30, 2]
		}
	];
	var dataChartPieLocs = [
		{
			name  : 'Location',
			labels: ['CN', 'DE', 'GB', 'MX', 'JP', 'IN', 'US'],
			values: [  69,   35,   40,   85,   38,   99,  101]
		}
	];

	var arrDataLineStat = [];
	{
		var tmpObjRed = { name:'Red', labels:QTRS, values:[] };
		var tmpObjAmb = { name:'Amb', labels:QTRS, values:[] };
		var tmpObjGrn = { name:'Grn', labels:QTRS, values:[] };
		var tmpObjUnk = { name:'Unk', labels:QTRS, values:[] };

		for (var idy=0; idy<QTRS.length; idy++) {
			tmpObjRed.values.push( Math.floor(Math.random() * 30) + 1 );
			tmpObjAmb.values.push( Math.floor(Math.random() * 50) + 1 );
			tmpObjGrn.values.push( Math.floor(Math.random() * 80) + 1 );
			tmpObjUnk.values.push( Math.floor(Math.random() * 10) + 1 );
		}

		arrDataLineStat.push( tmpObjRed );
		arrDataLineStat.push( tmpObjAmb );
		arrDataLineStat.push( tmpObjGrn );
		arrDataLineStat.push( tmpObjUnk );
	}

	// SLIDE 1: Bar Chart ------------------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Bar Chart', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		var arrDataRegions = [
			{
				name  : 'Region 1',
				labels: ['May', 'June', 'July', 'August'],
				values: [26, 53, 100, 75]
			},
			{
				name  : 'Region 2',
				labels: ['May', 'June', 'July', 'August'],
				values: [43.5, 70.3, 90.1, 80.05]
			}
		];
		var arrDataHighVals = [
			{
				name  : 'California',
				labels: ['Apartment', 'Townhome', 'Duplex', 'House', 'Big House'],
				values: [2000, 2800, 3200, 4000, 5000]
			},
			{
				name  : 'Texas',
				labels: ['Apartment', 'Townhome', 'Duplex', 'House', 'Big House'],
				values: [1400, 2000, 2500, 3000, 3800]
			}
		];

		// TOP-LEFT: H/bar
		var optsChartBar1 = { x:0.5, y:0.6, w:6.0, h:3.0,
			barDir: 'bar',
			border: { pt:'3', color:'00EE00' },
			fill: 'F1F1F1',

			catAxisLabelColor   : 'CC0000',
			catAxisLabelFontFace: 'Helvetica Neue',
			catAxisLabelFontSize: 14,
			catAxisOrientation  : 'maxMin',

			titleColor   : '33CF22',
			titleFontFace: 'Helvetica Neue',
			titleFontSize: 24
		};
		slide.addChart( pptx.charts.BAR, arrDataRegions, optsChartBar1 );

		// TOP-RIGHT: V/col
		var optsChartBar2 = { x:7.0, y:0.6, w:6.0, h:3.0,
			barDir: 'col',

			catAxisLabelColor   : '0000CC',
			catAxisLabelFontFace: 'Courier',
			catAxisLabelFontSize: 12,
			catAxisOrientation  : 'minMax',

			dataBorder         : { pt:'1', color:'F1F1F1' },
			dataLabelColor     : '696969',
			dataLabelFontFace  : 'Arial',
			dataLabelFontSize  : 11,
			dataLabelPosition  : 'outEnd',
			dataLabelFormatCode: '#.0',
			showValue          : true,

			valAxisOrientation: 'maxMin',

			showLegend: false,
			showTitle : false
		};
		slide.addChart( pptx.charts.BAR, arrDataRegions, optsChartBar2 );

		// BTM-LEFT: H/bar - TITLE and LEGEND
		slide.addText( '.', { x:0.5, y:3.8, w:6.0, h:3.5, fill:'F1F1F1', color:'F1F1F1'} );
		var optsChartBar3 = { x:0.5, y:3.8, w:6.0, h:3.5,
			barDir     : 'bar',

			border: { pt:'3', color:'CF0909' },
			fill: 'F1C1C1',

			catAxisLabelColor   : 'CC0000',
			catAxisLabelFontFace: 'Helvetica Neue',
			catAxisLabelFontSize: 14,
			catAxisOrientation  : 'minMax',

			titleColor   : '33CF22',
			titleFontFace: 'Helvetica Neue',
			titleFontSize: 16,

			showTitle : true,
			title: 'Sales by Region'
		};
		slide.addChart( pptx.charts.BAR, arrDataHighVals, optsChartBar3 );

		// BTM-RIGHT: V/col - TITLE and LEGEND
		slide.addText( '.', { x:7.0, y:3.8, w:6.0, h:3.5, fill:'F1F1F1', color:'F1F1F1'} );
		var optsChartBar4 = { x:7.0, y:3.8, w:6.0, h:3.5,
			barDir: 'col',
			barGapWidthPct: 25,
			chartColors: ['0088CC', '99FFCC'],
			chartColorsOpacity: 50,
			valAxisMaxVal: 5000,

			catAxisLabelColor   : '0000CC',
			catAxisLabelFontFace: 'Times',
			catAxisLabelFontSize: 11,
			catAxisOrientation  : 'minMax',

			dataBorder         : { pt:'1', color:'F1F1F1' },
			dataLabelColor     : 'FFFFFF',
			dataLabelFontFace  : 'Arial',
			dataLabelFontSize  : 10,
			dataLabelPosition  : 'ctr',
			showValue          : true,

			showLegend: true,
			legendPos :  't',
			showTitle: true,
			title    : 'Chart Title'
		};
		slide.addChart( pptx.charts.BAR, arrDataHighVals, optsChartBar4 );
	}

	// SLIDE: Date axis
	{
		var slide = pptx.addNewSlide();
		// These labels are Excel 1900 date system. Which is number of days after 1900 + 2
		var labels = [37987,38018,38047,38078,38108,38139,38169,38200,38231,38261,38292,38322,38353,38384,38412,38443,38473,38504,38534,38565,38596,38626,38657,38687,38718,38749,38777,38808,38838,38869,38899,38930,38961,38991,39022,39052,39083,39114,39142,39173,39203,39234,39264,39295,39326,39356,39387,39417,39448,39479,39508,39539,39569,39600,39630,39661,39692,39722,39753,39783,39814,39845,39873,39904,39934,39965,39995,40026,40057,40087,40118,40148,40179,40210,40238,40269,40299,40330,40360,40391,40422,40452,40483,40513,40544,40575,40603,40634,40664,40695,40725,40756,40787,40817,40848,40878,40909,40940,40969,41000,41030,41061,41091,41122,41153,41183,41214,41244,41275,41306,41334,41365,41395,41426,41456,41487,41518,41548,41579,41609,41640,41671,41699,41730,41760,41791,41821,41852,41883,41913,41944,41974,42005,42036,42064,42095,42125,42156,42186,42217,42248,42278,42309,42339,42370,42401,42430,42461,42491,42522,42552,42583,42614,42644,42675,42705,42736,42767,42795,42826,42856,42887,42917,42948];
		var data = [
            { name:'Beer', labels:labels, values: [41,42,43,42,43,42,44,44,42,41,40,41,37,39,39,39,41,41,44,45,42,43,40,41,37,39,39,38,40,42,43,44,43,39,36,41,36,37,37,37,37,38,39,40,37,38,35,37,33,33,35,35,37,37,39,39,35,35,34,35,33,34,36,36,36,38,42,43,40,38,35,36,34,35,35,35,37,39,43,43,39,38,36,40,37,38,37,39,40,43,45,43,41,40,37,42,38,39,40,42,45,46,46,48,46,41,40,44,39,42,43,41,45,47,49,49,45,43,42,45,42,44,46,46,47,50,52,52,45,44,42,47,43,43,45,48,50,51,54,56,50,47,45,49,44,45,47,49,50,53,54,50,51,48,46,50,46,48,50,52,52,57,59,58]},
            { name:'Wine', labels:labels, values: [75,73,70,70,69,70,70,74,72,71,80,100,72,74,70,68,71,69,71,75,72,75,81,98,71,69,64,66,66,68,68,71,71,70,78,97,69,70,66,66,64,65,67,67,67,67,73,88,63,63,61,61,61,61,62,62,62,60,67,79,60,60,57,57,57,55,58,59,58,59,64,75,57,56,53,53,55,54,58,59,60,59,64,77,59,59,54,55,56,56,59,57,57,58,63,75,58,58,53,55,55,56,56,56,59,56,62,75,56,57,56,54,55,55,55,58,56,57,64,76,56,58,55,56,57,54,56,57,56,57,64,76,57,58,55,55,56,55,57,57,59,59,66,79,60,61,56,58,61,62,60,58,63,62,67,83,62,62,59,61,61,61,64,63] },
            { name:'Soda', labels:labels, values: [null,null,null,12,12,11,11,11,11,11,13,11,11,12,13,11,12,11,11,11,12,12,12,11,11,12,13,12,12,11,11,11,11,11,11,11,11,11,12,11,11,16,14,13,14,21,18,15,12,12,13,12,11,11,11,12,12,12,11,12,12,12,13,12,12,12,12,12,12,12,12,12,12,12,13,13,13,12,13,13,13,13,13,13,13,13,14,14,14,14,16,15,14,14,15,16,15,15,16,15,16,17,17,17,17,16,17,18,17,17,19,17,17,17,18,18,17,16,17,18,19,20,20,18,18,17,18,18,21,18,21,25,25,23,25,24,25,29,27,26,25,24,25,26,26,25,26,25,24,24,25,24,24,24,24,25,26,25,28,27,25,25,27,26] }
        ];
        var opts = {
            x: 0.5,
            y: 0.5,
            w: 12,
            h: 6,
            showTitle: true,
            title: "Beverage Trends",
            showLegend: true,
            catLabelFormatCode: "mmm-yy",
            lineDataSymbol: "none"
        };
		slide.addChart(pptx.charts.LINE,
			data, opts);
	}

	// SLIDE 2: Bar Chart Grid/Axis Options ------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Bar Chart Grid/Axis Options', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		var arrDataRegions = [
			{
				name  : 'Region 1',
				labels: ['May', 'June', 'July', 'August'],
				values: [26, 53, 100, 75]
			},
			{
				name  : 'Region 2',
				labels: ['May', 'June', 'July', 'August'],
				values: [43.5, 70.3, 90.1, 80.05]
			}
		];
		var arrDataHighVals = [
			{
				name  : 'California',
				labels: ['Apartment', 'Townhome', 'Duplex', 'House', 'Big House'],
				values: [2000, 2800, 3200, 4000, 5000]
			},
			{
				name  : 'Texas',
				labels: ['Apartment', 'Townhome', 'Duplex', 'House', 'Big House'],
				values: [1400, 2000, 2500, 3000, 3800]
			}
		];

		// TOP-LEFT: H/bar
		var optsChartBar1 = { x:0.5, y:0.6, w:6.0, h:3.0,
			barDir: 'bar',
			fill: 'F1F1F1',

			catAxisLabelColor   : 'CC0000',
			catAxisLabelFontFace: 'Helvetica Neue',
			catAxisLabelFontSize: 14,

			catGridLine: 'none',
			catAxisHidden: true,
			valGridLine: { color: "cc6699", style: "dash", size: 1 },

			showLegend   : true,
			title        : 'No CatAxis, ValGridLine=dash',
			titleColor   : 'a9a9a9',
			titleFontFace: 'Helvetica Neue',
			titleFontSize: 14,
			showTitle    : true
		};
		slide.addChart( pptx.charts.BAR, arrDataRegions, optsChartBar1 );

		// TOP-RIGHT: V/col
		var optsChartBar2 = { x:7.0, y:0.6, w:6.0, h:3.0,
			barDir: 'col',
			fill: 'E1F1FF',

			dataBorder         : { pt:'1', color:'F1F1F1' },
			dataLabelColor     : '696969',
			dataLabelFontFace  : 'Arial',
			dataLabelFontSize  : 11,
			dataLabelPosition  : 'outEnd',
			dataLabelFormatCode: '#.0',
			showValue          : true,

			catAxisHidden: true,
			catGridLine  : 'none',
			valAxisHidden: true,
			valGridLine  : 'none',

			showLegend: true,
			legendPos : 'b',
			showTitle : false
		};
		slide.addChart( pptx.charts.BAR, arrDataRegions, optsChartBar2 );

		// BTM-LEFT: H/bar - TITLE and LEGEND
		slide.addText( '.', { x:0.5, y:3.8, w:6.0, h:3.5, fill:'F1F1F1', color:'F1F1F1'} );
		var optsChartBar3 = { x:0.5, y:3.8, w:6.0, h:3.5,
			barDir     : 'bar',

			border: { pt:'3', color:'CF0909' },
			fill: 'F1C1C1',

			catAxisLabelColor   : 'CC0000',
			catAxisLabelFontFace: 'Helvetica Neue',
			catAxisLabelFontSize: 14,
			catAxisOrientation  : 'maxMin',
			catAxisTitle: "Housing Type",
			catAxisTitleColor: "428442",
			catAxisTitleFontSize: 14,
			showCatAxisTitle: true,

			valGridLine: 'none',
			valAxisHidden: true,
			catGridLine: { color: "cc6699", style: "dash", size: 1 },

			titleColor   : '33CF22',
			titleFontFace: 'Helvetica Neue',
			titleFontSize: 16,

			showTitle : true,
			title: 'Sales by Region'
		};
		slide.addChart( pptx.charts.BAR, arrDataHighVals, optsChartBar3 );

		// BTM-RIGHT: V/col - TITLE and LEGEND
		slide.addText( '.', { x:7.0, y:3.8, w:6.0, h:3.5, fill:'F1F1F1', color:'F1F1F1'} );
		var optsChartBar4 = { x:7.0, y:3.8, w:6.0, h:3.5,
			barDir: 'col',
			barGapWidthPct: 25,
			chartColors: ['0088CC', '99FFCC'],
			chartColorsOpacity: 50,
			valAxisMinVal: 1000,
			valAxisMaxVal: 5000,

			catAxisLabelColor   : '0000CC',
			catAxisLabelFontFace: 'Times',
			catAxisLabelFontSize: 11,
			catAxisOrientation  : 'minMax',

			dataBorder         : { pt:'1', color:'F1F1F1' },
			dataLabelColor     : 'FFFFFF',
			dataLabelFontFace  : 'Arial',
			dataLabelFontSize  : 10,
			dataLabelPosition  : 'ctr',
			showValue          : true,

			valAxisHidden      : true,
			catAxisTitle       : 'Housing Type',
			showCatAxisTitle   : true,

			showLegend: false,
			showTitle : false
		};
		slide.addChart( pptx.charts.BAR, arrDataHighVals, optsChartBar4 );
	}

	// SLIDE 3: Stacked Bar Chart ----------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Bar Chart: Stacked/PercentStacked', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		var arrDataRegions = [
			{
				name  : 'Region 3',
				labels: ['April', 'May', 'June', 'July', 'August'],
				values: [17, 26, 53, 100, 75]
			},
			{
				name  : 'Region 4',
				labels: ['April', 'May', 'June', 'July', 'August'],
				values: [55, 43, 70, 90, 80]
			}
		];
		var arrDataHighVals = [
			{
				name  : 'California',
				labels: ['Apartment', 'Townhome', 'Duplex', 'House', 'Big House'],
				values: [2000, 2800, 3200, 4000, 5000]
			},
			{
				name  : 'Texas',
				labels: ['Apartment', 'Townhome', 'Duplex', 'House', 'Big House'],
				values: [1400, 2000, 2500, 3000, 3800]
			}
		];

		// TOP-LEFT: H/bar
		var optsChartBar1 = { x:0.5, y:0.6, w:6.0, h:3.0,
			barDir: 'bar',
			barGrouping: 'stacked',

			catAxisLabelColor   : 'CC0000',
			catAxisLabelFontFace: 'Helvetica Neue',
			catAxisLabelFontSize: 14,
			catAxisOrientation  : 'maxMin',

			dataLabelColor   : 'FFFFFF',
			showValue        : true,

			titleColor   : '33CF22',
			titleFontFace: 'Helvetica Neue',
			titleFontSize: 24
		};
		slide.addChart( pptx.charts.BAR, arrDataRegions, optsChartBar1 );

		// TOP-RIGHT: V/col
		var optsChartBar2 = { x:7.0, y:0.6, w:6.0, h:3.0,
			barDir: 'col',
			barGrouping: 'stacked',

			dataLabelColor   : 'FFFFFF',
			dataLabelFontFace: 'Arial',
			dataLabelFontSize: 12,
			showValue        : true,

			catAxisLabelColor   : '0000CC',
			catAxisLabelFontFace: 'Courier',
			catAxisLabelFontSize: 12,
			catAxisOrientation  : 'minMax',

			showLegend: false,
			showTitle : false
		};
		slide.addChart( pptx.charts.BAR, arrDataRegions, optsChartBar2 );

		// BTM-LEFT: H/bar - 100% layout without axis labels
		var optsChartBar3 = { x:0.5, y:3.8, w:6.0, h:3.5,
			barDir       : 'bar',
			barGrouping  : 'percentStacked',
			dataBorder   : { pt:'1', color:'F1F1F1' },
			catAxisHidden: true,
			valAxisHidden: true,
			showTitle    : false,
			layout       : {x:0.1, y:0.1, w:1, h:1}
		};
		slide.addChart( pptx.charts.BAR, arrDataRegions, optsChartBar3 );

		// BTM-RIGHT: V/col - TITLE and LEGEND
		slide.addText( '.', { x:7.0, y:3.8, w:6.0, h:3.5, fill:'F1F1F1', color:'F1F1F1'} );
		var optsChartBar4 = { x:7.0, y:3.8, w:6.0, h:3.5,
			barDir: 'col',
			barGrouping: 'percentStacked',

			catAxisLabelColor   : '0000CC',
			catAxisLabelFontFace: 'Times',
			catAxisLabelFontSize: 12,
			catAxisOrientation  : 'minMax',

			chartColors: ['5DA5DA','FAA43A'],

			showLegend: true,
			legendPos :  't'
		};
		slide.addChart( pptx.charts.BAR, arrDataHighVals, optsChartBar4 );
	}

	// SLIDE 4: Bar Chart - Lots of Bars ---------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Lots of Bars (>26 letters)', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		var arrDataHighVals = [
			{
				name  : 'TEST: getExcelColName',
				labels: LETTERS.concat(['AA','AB','AC','AD']),
				values: [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30 ]
			}
		];

		var optsChart = {
			x:0.5, y:0.5, w:'90%', h:'90%',
			barDir: 'col',
			title: 'Chart With >26 Cols',
			showTitle: true,
			titleFontSize: 20,
			titleRotate: 10,
			showCatAxisTitle: true,
			catAxisTitle: "Letters",
			catAxisTitleColor: "4286f4",
			catAxisTitleFontSize: 14,

			showValAxisTitle: true,
			valAxisTitle: "Column Index",
			valAxisTitleColor: "c11c13",
			valAxisTitleFontSize: 16,
		};

		// TEST `getExcelColName()` to ensure Excel Column names are generated correctly above >26 chars/cols
		slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChart);
	}

	// SLIDE 5: Bar Chart: Data Series Colors, majorUnits, and valAxisLabelFormatCode ------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Bar Colors, valAxisMajorUnit, v Format %', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		// TOP-LEFT
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name  : 'Colored Series',
					labels: ['Jan', 'Feb','Mar', 'Apr', 'May', 'Jun'],
					values: [20, 30, 10, 25, 15, 5]
				}
			],
			{ x:0.5, y:0.6, w:'45%', h:3,
				valAxisMaxVal:100,
				barDir: 'bar',
				valAxisMajorUnit: 50,
				chartColors: ['0077BF','4E9D2D','ECAA00','5FC4E3','DE4216','154384'],
				barGapWidthPct: 25
			}
		);

		// TOP-RIGHT
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name  : 'Too Many Colors Series',
					labels: ['Jan', 'Feb','Mar', 'Apr', 'May', 'Jun'],
					values: [.20, .30, .10, .25, .15, .05]
				}
			],
			{ x:7, y:0.6, w:'45%', h:3,
				valAxisMaxVal:1,
				barDir: 'bar',
				showValue: true,
				dataLabelPosition: 'outEnd',
				dataLabelFormatCode: '#%',
				valAxisLabelFormatCode: '#%',
				valAxisMajorUnit: 0.2,
				chartColors: ['0077BF','4E9D2D','ECAA00','5FC4E3','DE4216','154384', '7D666A','A3C961','EF907B','9BA0A3'],
				barGapWidthPct: 25
			}
		);

		// BOTTOM-LEFT
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name  : 'Two Color Series',
					labels: ['Jan', 'Feb','Mar', 'Apr', 'May', 'Jun'],
					values: [.20, .30, .10, .25, .15, .05]
				}
			],
			{  x:0.5, y:3.6, w:'45%', h:3,
				valAxisMaxVal:1,
				barDir: 'bar',
				showValue: true,
				dataLabelPosition: 'outEnd',
				dataLabelFormatCode: '#%',
				valAxisLabelFormatCode: '0.#0',
				chartColors: ['DE4216','154384'],
				barGapWidthPct: 25
			}
		);

		// BOTTOM-RIGHT
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name  : 'Many values',
					labels: 'A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P'.split(','),
					values: [1.20, 2.30, 3.10, 4.25, 2.15, 6.05, 8.01, 2.02, 9.9, 0.9, 2.2, 3.33, 6.66, 7.77, 8.88]
				}
			], {  x:7, y:3.6, w:'45%', h:3,
				valAxisMaxVal:10,
				barDir: 'bar',
				showValue: true,
				dataLabelPosition: 'outEnd',
				chartColors: ['0077BF','4E9D2D','ECAA00','5FC4E3','DE4216','154384', '7D666A','A3C961','EF907B','9BA0A3'],
				barGapWidthPct: 25,
				valAxisMajorUnit: 1
			}
		);
	}

	// SLIDE 6: Tornado Chart -------------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Tornado Chart - Grid and Axis Formatting', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name: 'High',
					labels: ['London', 'Munich', 'Tokyo'],
					values: [.20, .32, .41]
				},
				{
					name: 'Low',
					labels: ['London', 'Munich', 'Tokyo'],
					values: [-0.11, -0.22, -0.29]
				}
			],
			{
				x:0.5, y:0.5, w:'90%', h:'90%',
				valAxisMaxVal:1,
				barDir: 'bar',
				axisLabelFormatCode: '#%',
				gridLineColor: 'D8D8D8',
				axisLineColor: 'D8D8D8',
				catAxisLineShow: false,
				valAxisLineShow: false,
				barGrouping: 'stacked',
				catAxisLabelPos: 'low',
				valueBarColors: true,
				//dataLabelPosition: 'outEnd',
				chartColors: ['0077BF','4E9D2D','ECAA00','5FC4E3','DE4216','154384','7D666A','A3C961','EF907B','9BA0A3'],
				invertedColors: ['0065A2','428526','C99100','51A7C1','BD3813','123970','6A575A','8BAB52','CB7A69','84888B'],
				barGapWidthPct: 25,
				valAxisMajorUnit: 0.2
			}
		);
	}

	// SLIDE 7: Line Chart: Line Smoothing, Line Size, Symbol Size -------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Line Smoothing, Line Size, Line Shadow, Symbol Size', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		slide.addText( '..', { x:0.5, y:0.6, w:6.0, h:3.0, fill:'F1F1F1', color:'F1F1F1'} );
		var optsChartLine1 = { x:0.5, y:0.6, w:6.0, h:3.0,
			chartColors: [ COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK ],
			lineSize  : 8,
			lineSmooth: true,
			showLegend: true, legendPos: 't'
		};
		slide.addChart( pptx.charts.LINE, arrDataLineStat, optsChartLine1 );

		var optsChartLine2 = { x:7.0, y:0.6, w:6.0, h:3.0,
			chartColors: [ COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK ],
			lineSize  : 16,
			lineSmooth: true,
			showLegend: true, legendPos: 'r'
		};
		slide.addChart( pptx.charts.LINE, arrDataLineStat, optsChartLine2 );

		var optsChartLine1 = { x:0.5, y:4.0, w:6.0, h:3.0,
			chartColors: [ COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK ],
			lineDataSymbolSize: 10,
			lineShadow: 'none',
			showLegend: true, legendPos: 'l'
		};
		slide.addChart( pptx.charts.LINE, arrDataLineStat, optsChartLine1 );

		var shadowOpts = { type:'outer', color:'cd0011', blur:3, offset:12, angle:75, opacity:0.8 };
		var optsChartLine2 = { x:7.0, y:4.0, w:6.0, h:3.0,
			chartColors: [ COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK ],
			lineDataSymbolSize: 20,
			lineShadow: shadowOpts,
			showLegend: true, legendPos: 'b'
		};
		slide.addChart( pptx.charts.LINE, arrDataLineStat, optsChartLine2 );
	}

	// SLIDE 8: Line Chart: TEST: `lineDataSymbol` + `lineDataSymbolSize` ------------------
	{
		var intWgap = 4.25;
		var opts_lineDataSymbol = ['circle','dash','diamond','dot','none','square','triangle'];
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Line Chart: lineDataSymbol option test', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		opts_lineDataSymbol.forEach(function(opt,idx){
			slide.addChart(
				pptx.charts.LINE,
				arrDataLineStat,
				{
					x:(idx < 3 ? idx*intWgap : (idx < 6 ? (idx-3)*intWgap : (idx-6)*intWgap)), y:(idx < 3 ? 0.5 : (idx < 6 ? 2.75 : 5)),
					w:4.25, h:2.25,
					lineDataSymbol:opt, title:opt, showTitle:true,
					lineDataSymbolSize:(idx==5 ? 9 : (idx==6 ? 12 : null))
				}
			);
		});
	}

	// SLIDE 9: Line Chart: Lots of Cats ---------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Line Chart: Lots of Lines', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		var MAXVAL = 20000;

		var arrDataTimeline = [];
		for (var idx=0; idx<15; idx++) {
			var tmpObj = {
				name  : 'Series'+idx,
				labels: MONS,
				values: []
			};

			for (var idy=0; idy<MONS.length; idy++) {
				tmpObj.values.push( Math.floor(Math.random() * MAXVAL) + 1 );
			}

			arrDataTimeline.push( tmpObj );
		}

		// FULL SLIDE:
		var optsChartLine1 = { x:0.5, y:0.6, w:'95%', h:'85%',
			fill: 'F2F9FC',

			valAxisMaxVal: MAXVAL,

			showLegend: true,
			legendPos : 'r'
		};
		slide.addChart( pptx.charts.LINE, arrDataTimeline, optsChartLine1 );
	}

	// SLIDE 10: Area Chart: Misc -----------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Area Chart', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		var arrDataTimeline2ser = [
			{
				name  : 'Actual Sales',
				labels: MONS,
				values: [1500, 4600, 5156, 3167, 8510, 8009, 6006, 7855, 12102, 12789, 10123, 15121]
			},
			{
				name  : 'Proj Sales',
				labels: MONS,
				values: [1000, 2600, 3456, 4567, 5010, 6009, 7006, 8855, 9102, 10789, 11123, 12121]
			}
		];

		// TOP-LEFT
		var optsChartLine1 = { x:0.5, y:0.6, w:'45%', h:3, catAxisOrientation:'maxMin', valAxisOrientation:'maxMin' };
		slide.addChart( pptx.charts.AREA, arrDataTimeline2ser, optsChartLine1 );

		// TOP-RIGHT
		var optsChartLine2 = { x:7, y:0.6, w:'45%', h:3,
			chartColors: ['0088CC', '99FFCC'],
			chartColorsOpacity: 25,
			dataBorder: {pt:2, color:'FFFFFF'},
			fill: 'D1E1F1'
		};
		slide.addChart( pptx.charts.AREA, arrDataTimeline2ser, optsChartLine2 );

		// BOTTOM-LEFT
		var optsChartLine3 = { x:0.5, y:3.6, w:'45%', h:3,
			chartColors: ['0088CC', '99FFCC'],
			chartColorsOpacity: 50
		};
		slide.addChart( pptx.charts.AREA, arrDataTimeline2ser, optsChartLine3 );

		// BOTTOM-RIGHT
		var optsChartLine4 = { x:7, y:3.6, w:'45%', h:3,
			chartColors: ['CC8833', 'CCFF69'],
			chartColorsOpacity: 75
		};
		slide.addChart( pptx.charts.AREA, arrDataTimeline2ser, optsChartLine4 );
	}

	// SLIDE 11: Pie Charts: All 4 Legend Options ------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Pie Charts: Legends', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		// INTERNAL USE: Not visible to user (its behind a chart): Used for ensuring ref counting works across obj types (eg: `rId` check/test)
		slide.addImage({ path:'images/cc_copyremix.gif', x:0.5, y:1.0, w:1.2, h:1.2 });

		// TOP-LEFT
		slide.addText( '.', {x:0.5, y:0.5, w:4.2, h:3.2, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart( pptx.charts.PIE, dataChartPieStat, {x:0.5, y:0.5, w:4.2, h:3.2, showLegend:true, legendPos:'l'} );

		// TOP-RIGHT
		slide.addText( '.', {x:5.6, y:0.5, w:3.2, h:3.2, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart( pptx.charts.PIE, dataChartPieStat, {x:5.6, y:0.5, w:3.2, h:3.2, showLegend:true, legendPos:'t'} );

		// BTM-LEFT
		slide.addText( '.', {x:0.5, y:4.0, w:4.2, h:3.2, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart( pptx.charts.PIE, dataChartPieLocs, {x:0.5, y:4.0, w:4.2, h:3.2, showLegend:true, legendPos:'r'} );

		// BTM-RIGHT
		slide.addText( '.', {x:5.6, y:4.0, w:3.2, h:3.2, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart( pptx.charts.PIE, dataChartPieLocs, {x:5.6, y:4.0, w:3.2, h:3.2, showLegend:true, legendPos:'b'} );

		// BOTH: TOP-RIGHT
		slide.addText( '.', {x:9.8, y:0.5, w:3.2, h:3.2, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart( pptx.charts.PIE, dataChartPieLocs, {x:9.8, y:0.5, w:3.2, h:3.2, dataBorder:{pt:'1',color:'F1F1F1'}, showLegend:true, legendPos:'t', showTitle:true, title:'Title & Legend'} );

		// BOTH: BTM-RIGHT
		slide.addText( '.', {x:9.8, y:4.0, w:3.2, h:3.2, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart( pptx.charts.PIE, dataChartPieLocs, {x:9.8, y:4.0, w:3.2, h:3.2, dataBorder:{pt:'1',color:'F1F1F1'}, showLegend:true, legendPos:'b', showTitle:true, title:'Title & Legend'} );
	}

	// SLIDE 12: Doughnut Chart ------------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Chart Examples: Doughnut Chart', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		var optsChartPie1 = {
			x:0.5, y:1.0, w:6.0, h:6.0,
			chartColors: ['FC0000','FFCC00','009900','6600CC'],
			dataBorder       : { pt:'2', color:'F1F1F1' },
			dataLabelColor   : 'FFFFFF',
			dataLabelFontSize: 14,

			legendPos : 'r',

			showLabel  : false,
			showValue  : false,
			showPercent: true,
			showLegend : true,
			showTitle  : false,

			holeSize: 70,

			title        : 'Project Status',
			titleColor   : '33CF22',
			titleFontFace: 'Helvetica Neue',
			titleFontSize: 24
		};
		slide.addText( '.', {x:0.5, y:1.0, w:6.0, h:6.0, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart(pptx.charts.DOUGHNUT, dataChartPieStat, optsChartPie1 );

		var optsChartPie2 = {
			x:7.0, y:1.0, w:6, h:6,
			dataBorder       : { pt:'3', color:'F1F1F1' },
			dataLabelColor   : 'FFFFFF',
			showLabel  : true,
			showValue  : true,
			showPercent: true,
			showLegend : false,
			showTitle  : false,
			title: 'Resource Totals by Location'
		};
		slide.addChart(pptx.charts.DOUGHNUT, dataChartPieLocs, optsChartPie2 );
	}
}

function genSlides_Media(pptx) {
	// SLIDE 1: Video and YouTube
	// ======== -----------------------------------------------------------------------------------
	var slide1 = pptx.addNewSlide();
	slide1.addTable( [ [{ text:'Media: Video Examples', opts:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

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
	slide2.addTable( [ [{ text:'Media: Audio and Pre-Encoded Audio/Video Examples', opts:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

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
	slide.addTable( [ [{ text:'Image Examples: Misc Image Types', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

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

	// Image Link
	slide.addText('Hyperlink Image', { x:9.6, y:0.6, w:3.0, h:0.4, color:'0088CC' });
	slide.addImage({
		x:9.8, y:1.2, w:0.8, h:0.8,
		hyperlink:{ url:'https://github.com/gitbrent/pptxgenjs', tooltip:'Visit Homepage' },
		data:
			'image/jpeg;base64,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'
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
	slide.addTable( [ [{ text:'Shape Examples 1: Misc Shape Types (no text)', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

	//slide.addShape(pptx.shapes.RECTANGLE,         { x:0.5, y:0.8, w:12.5,h:0.5, fill:'F9F9F9' });
	slide.addShape(pptx.shapes.RECTANGLE,         { x:0.5, y:0.8, w:1.5, h:3.0, fill:'FF0000' });
	slide.addShape(pptx.shapes.RECTANGLE,         { x:3.0, y:0.7, w:1.5, h:3.0, fill:'F38E00', rotate:45 });
	slide.addShape(pptx.shapes.OVAL,              { x:5.4, y:0.8, w:3.0, h:1.5, fill:{ type:'solid', color:'0088CC', alpha:25 } });
	slide.addShape(pptx.shapes.OVAL,              { x:7.7, y:1.4, w:3.0, h:1.5, fill:{ type:'solid', color:'FF00CC', alpha:50 }, rotate:90 });
	slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:10 , y:2.5, w:3.0, h:1.5, r:0.2, fill:'00FF00', line:'000000', line_size:1 });
	//
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:4.4, w:5.0, h:0.0, line:'FF0000', line_size:1 });
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:4.8, w:5.0, h:0.0, line:'FF0000', line_size:2, line_head:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:5.2, w:5.0, h:0.0, line:'FF0000', line_size:3, line_tail:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:5.6, w:5.0, h:0.0, line:'FF0000', line_size:4, line_head:'triangle', line_tail:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:5.7, y:3.3, w:2.5, rotate:(360-45) }); // DIAGONAL Line // TEST: (missing `h`, `line`, `line_size`)
	//
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:0.4, y:4.3, w:6.0, h:3.0, fill:'0088CC', line:'000000', line_size:3 });
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:7.0, y:4.3, w:6.0, h:3.0, fill:'0088CC', line:'000000', flipH:true });

	// SLIDE 2: Misc Shape Types with Text
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Shape Examples 2: Misc Shape Types (with text)', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

	slide.addText('RECTANGLE',                  { shape:pptx.shapes.RECTANGLE,         x:0.5, y:0.8, w:1.5, h:3.0, fill:'FF0000', align:'c', font_size:14 });
	slide.addText('RECTANGLE (rotate:45)',      { shape:pptx.shapes.RECTANGLE,         x:3.0, y:0.7, w:1.5, h:3.0, fill:'F38E00', rotate:45, align:'c', font_size:14 });
	slide.addText('OVAL (alpha:25)',            { shape:pptx.shapes.OVAL,              x:5.4, y:0.8, w:3.0, h:1.5, fill:{ type:'solid', color:'0088CC', alpha:25 }, align:'c', font_size:14 });
	slide.addText('OVAL (rotate:90, alpha:50)', { shape:pptx.shapes.OVAL,              x:7.7, y:1.4, w:3.0, h:1.5, fill:{ type:'solid', color:'FF00CC', alpha:50 }, rotate:90, align:'c', font_size:14 });
	slide.addText('ROUNDED-RECTANGLE\nline_dash:dash\nrectRadius:10', { shape:pptx.shapes.ROUNDED_RECTANGLE, x:10 , y:2.5, w:3.0, h:1.5, r:0.2, fill:'00FF00', align:'c', font_size:14, line:'000000', line_size:1, line_dash:'dash', rectRadius:10 });
	//
	slide.addText('LINE size=1',     { shape:pptx.shapes.LINE, align:'c', x:4.15, y:4.40, w:5, h:0, line:'FF0000', line_size:1, line_dash:'lgDash' });
	slide.addText('LINE size=2',     { shape:pptx.shapes.LINE, align:'l', x:4.15, y:4.80, w:5, h:0, line:'FF0000', line_size:2, line_tail:'triangle' });
	slide.addText('LINE size=3',     { shape:pptx.shapes.LINE, align:'r', x:4.15, y:5.20, w:5, h:0, line:'FF0000', line_size:3, line_head:'triangle' });
	slide.addText('LINE size=4',     { shape:pptx.shapes.LINE, align:'c', x:4.15, y:5.60, w:5, h:0, line:'FF0000', line_size:4, line_head:'triangle', line_tail:'triangle' });
	slide.addText('DIAGONAL',        { shape:pptx.shapes.LINE, valign:'b', x:5.7, y:3.3, w:2.5, line_size:2, rotate:(360-45) }); // TEST: (missing `h` and `line`)
	//
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:0.4, y:4.3, w:6, h:3, fill:'0088CC', line:'000000', line_size:3 });
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:7.0, y:4.3, w:6, h:3, fill:'0088CC', line:'000000', flipH:true });
}

function genSlides_Text(pptx) {
	// SLIDE 1: Line Break / Bullets
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Text Examples 1', options:gOptsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

		// LEFT COLUMN ------------------------------------------------------------

		// 1: Multi-Line Formatting
		slide.addText("Word-Level Formatting:", { x:0.5, y:0.5, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text:'1st\nline',options:{ font_size:24, font_face:'Courier New', color:'99ABCC', align:'r', breakLine:true } },
				{ text:'2nd line', options:{ font_size:36, font_face:'Arial',       color:'FFFF00', align:'c', breakLine:true } },
				{ text:'3rd line', options:{ font_size:48, font_face:'Verdana',     color:'0088CC', align:'l' } },
				{ text:'4th line', options:{ font_size:38, font_face:'Arial',       color:'FFFF00', align:'c' } },
				{ text:'5th\nline',options:{ font_size:36, font_face:'Courier New', color:'99ABCC', align:'r' } }
			],
			{ x:0.5, y:0.85, w:6, h:4, margin:0.1, fill:'232323' }
		);

		// 3: Hyperlinks
		slide.addText("Hyperlinks:", { x:0.5, y:5.0, w:1.75, h:0.35, color:'0088CC' });
		slide.addText(
			[
				{ text:'Visit the ' },
				{ text:'PptxGenJS Project', options:{ hyperlink:{ url:'https://github.com/gitbrent/pptxgenjs', tooltip:'Visit Homepage' } } },
				{ text:' or ' },
				{ text:'(link without tooltip)', options:{hyperlink:{url:'https://github.com/gitbrent'}} }
			],
			{ x:0.5, y:5.35, w:6.0, h:0.6, margin:0.1, fill:'F1F1F1', font_size:14 }
		);

		// 4: Text Effects: Shadow
		var shadowOpts = { type:'outer', color:'696969', blur:3, offset:10, angle:45, opacity:0.8 };
		slide.addText("Text Shadow:", { x:0.5, y:6.2, w:'50%', h:0.38, color:'0088CC' });
		slide.addText(
			'Outer Shadow (blur:3, offset:10, angle:45, opacity:80%)',
			{ x:0.5, y:6.5, w:12, h:0.6, font_size:32, color:'0088cc', shadow:shadowOpts }
		);

		// RIGHT COLUMN ------------------------------------------------------------

		// 2: Line-Break Test
		slide.addText("Line-Breaks:", { x:7.5, y:0.5, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			'***Line-Break/Multi-Line Test***\n\nFirst line\nSecond line\nThird line',
			{ x:7.5, y:0.85, w:5.25, h:1.6, valign:'middle', align:'ctr', color:'6c6c6c', font_size:16, fill:'F2F2F2' }
		);

		slide.addText("Line-Spacing (text):", { x:7.5, y:2.6, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			'1st Line\n2nd Line (40pt)',
			{ x:7.5, y:2.95, w:5.25, h:1.25, valign:'m', align:'c', fill:'f1f1f1', color:'363636', lineSpacing:40 }
		);

		slide.addText("Line-Spacing (bullets):", { x:7.5, y:4.45, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
			  { text:'Line Spacing\n35pt', options:{ font_size:24, bullet:true, color:'99ABCC', lineSpacing:35 } }
			],
			{ x:7.5, y:4.85, w:5.25, h:1.15, margin:0.1, fill:'f1f1f1' }
		);
	}

	// SLIDE 2: Bullets
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Text Examples 2: Bullets', options:gOptsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

		// LEFT COLUMN ------------------------------------------------------------

		// 1: Bullets with indent levels
		slide.addText("Bullet Indent-Levels:", { x:0.5, y:0.5, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text:'Bullet one',     options:{ font_size:24, bullet:true, color:'99ABCC'                } },
				{ text:'Indent-Level 1', options:{ font_size:32, bullet:true, color:'FFFF00', indentLevel:1 } },
				{ text:'Indent-Level 2', options:{ font_size:42, bullet:true, color:'0088CC', indentLevel:2 } },
				{ text:'Indent-Level 3', options:{ font_size:48, bullet:true, color:'CC88BB', indentLevel:3 } },
				{ text:'Indent-Level 3', options:{ font_size:48, bullet:true, color:'CC88BB', indentLevel:3 } },
				{ text:'Indent-Level 2', options:{ font_size:42, bullet:true, color:'0088CC', indentLevel:2 } },
				{ text:'Indent-Level 2', options:{ font_size:42, bullet:true, color:'0088CC', indentLevel:2 } },
				{ text:'Indent-Level 1', options:{ font_size:32, bullet:true, color:'FFFF00', indentLevel:1 } },
				{ text:'Bullet no indent', options:{ font_size:24, bullet:true, color:'99ABCC'                } },
				{ text:'Bullet Last',    options:{ font_size:24, bullet:true, color:'99ABCC'                } }
			],
			{ x:0.5, y:1.0, w:6.25, h:6.0, fill:'373737' }
		);

		// 4: Regular bullets
		slide.addText("Bullets:", { x:7.5, y:0.65, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(12345                  , { x:8.0, y:1.05, w:'30%', h:0.5, color:'0000DE', font_face:"Courier New", bullet:true });
		slide.addText('String (number above)', { x:8.0, y:1.35, w:'30%', h:0.5, color:'00AA00', bullet:true });

		// 5: Bullets: Text With Line-Breaks
		slide.addText("Bullets with line-breaks:", { x:7.5, y:2.1, w:'40%', h:0.38, color:'0088CC' });
		slide.addText('Line 1\nLine 2\nLine 3', { x:8.0, y:2.6, w:'30%', h:1.0, color:'393939', font_size:16, fill:'F2F2F2', bullet:{type:'number'} });

		// 6: Bullets: With group of {text}
		slide.addText("Bullet with {text} objects:", { x:7.5, y:4.0, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text: 'big red words... ', options:{font_size:24, color:'FF0000'} },
				{ text: 'some green words.', options:{font_size:16, color:'00FF00'} }
			],
			{ x:8.0, y:4.4, w:5.0, h:0.5, margin:0.1, font_face:'Arial', bullet:{code:'25BA'} }
		);

		// 7: Bullets: Within a {text} object
		slide.addText("Bullet within {text} objects:", { x:7.5, y:5.3, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text:'I am a text object with bullets..', options:{bullet:{code:'2605'}, color:'CC0000'} },
				{ text:'and I am the next text object.'   , options:{bullet:{code:'25BA'}, color:'00CD00'} },
				{ text:'Default bullet text.. '           , options:{bullet:true, color:'696969'} },
				{ text:'Final text object w/ bullet:true.', options:{bullet:true, color:'0000AB'} }
			],
			{ x:8.0, y:5.65, w:'35%', h:1.4, color:'ABABAB', margin:1 }
		);
	}

	// SLIDE 3: Text alignment, percent x/y, etc.
	{
		var slide = pptx.addNewSlide();
		// Slide colors: bkgd/fore
		slide.back = '030303';
		slide.color = '9F9F9F';
		// Title
		slide.addTable( [ [{ text:'Text Examples 2', options:gOptsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

		// Actual Textbox shape (can have any Height, can wrap text, etc.)
		slide.addText( 'Textbox (ctr/ctr)', { x:0.5, y:0.75, w:8.5, h:2.5, color:'FFFFFF', fill:'0000FF', valign:'c', align:'c', isTextBox:true } );
		slide.addText( 'Textbox (top/lft)', { x:10,  y:0.75, w:3.0, h:1.0, color:'FFFFFF', fill:'00CC00', valign:'t', align:'l', isTextBox:true } );
		slide.addText( 'Textbox (btm/rgt)', { x:10,  y:2.25, w:3.0, h:1.0, color:'FFFFFF', fill:'FF0000', valign:'b', align:'r', isTextBox:true } );

		slide.addText('^ (50%/50%)', {x:'50%', y:'50%', w:2});

		slide.addText('Plain x/y coords', { x:10, y:3.5 });

		slide.addText('Escaped chars: \' " & < >', { x:10, y:4.5 });

		slide.addText(
			[
				{ text:'Sub'},
				{ text:'Subscript', options:{ subscript:true } },
				{ text:' // Super'},
				{ text:'Superscript', options:{ superscript:true } }
			],
			{ x:10, y:5.5 }
		);

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

	var slide4 = pptx.addNewSlide( pptx.masters.TITLE_SLIDE,  { bkgd:'0088CC', slideNumber:{x:'50%', y:'90%', color:'0088CC'} } );
	var slide5 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:{ path:'images/title_bkgd_alt.jpg' } } );
	var slide6 = pptx.addNewSlide( pptx.masters.THANKS_SLIDE, { bkgd:'ffab33'} );

	//var slide7 = pptx.addNewSlide( pptx.masters.LEGACY_TEST_ONLY );
}

// ==================================================================================================================

if ( typeof module !== 'undefined' && module.exports ) {
	module.exports = runEveryTest;
}
