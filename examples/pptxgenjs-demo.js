/**
* NAME: pptxgenjs-demo.js
* AUTH: Brent Ely (https://github.com/gitbrent/)
* DESC: Common test/demo slides for all library features
* DEPS: Loaded by `pptxgenjs-demo.js` and `nodejs-demo.js`
* VER.: 2.5.0-beta
* BLD.: 20190206
*/

// Detect Node.js (NODEJS is ultimately used to determine how to save: either `fs` or web-based, so using fs-detection is perfect)
var NODEJS = false;
{
	// NOTE: `NODEJS` determines which network library to use, so using fs-detection is apropos.
	if ( typeof module !== 'undefined' && module.exports && typeof require === 'function' && typeof window === 'undefined' ) {
		try {
			require.resolve('fs');
			NODEJS = true;
		}
		catch (ex) {
			NODEJS = false;
		}
	}
}
if (NODEJS) { var LOGO_STARLABS; }

// Constants
var CUST_NAME = 'S.T.A.R. Laboratories';
var USER_NAME = 'Barry Allen';
var COLOR_RED = 'FF0000';
var COLOR_AMB = 'F2AF00';
var COLOR_GRN = '7AB800';
var COLOR_CRT = 'AA0000';
var COLOR_BLU = '0088CC';
var COLOR_UNK = 'A9A9A9';
var ARRSTRBITES = [130];
var CHARSPERLINE = 130; // "Open Sans", 13px, 900px-colW = ~19 words/line ~130 chars/line
// Lorem text / base64 images
{
	// FYI: 3086 chars
	var gStrLoremIpsum =
		'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin condimentum dignissim velit vel luctus. Donec feugiat ipsum quis tempus blandit. Donec mattis mauris vel est dictum interdum. Pellentesque imperdiet nibh vitae porta ornare. Fusce non nisl lacus. Curabitur ut mattis dui. Ut pulvinar urna velit, vitae aliquam neque pulvinar eu. Fusce eget tellus eu lorem finibus mattis. Nunc blandit consequat arcu. Ut sed pharetra tortor, nec finibus ipsum. Pellentesque a est vitae ligula imperdiet rhoncus. Ut quis hendrerit tellus. Phasellus non malesuada mi. Suspendisse ullamcorper tristique odio fermentum elementum. Phasellus mattis mollis mauris, non mattis ligula dapibus quis. Quisque pretium metus massa. Curabitur condimentum consequat felis, id rutrum velit cursus vel. Proin nulla est, posuere in velit at, faucibus dignissim diam. Quisque quis erat euismod, malesuada erat eu, congue nisi. Ut risus lectus, auctor at libero sit amet, accumsan ultricies est. Donec eget iaculis enim. Nunc ac egestas tellus, nec efficitur magna. Sed nec nisl ut augue laoreet sollicitudin vitae nec quam. Vestibulum pretium nisl bibendum, tempor velit eu, semper velit. Nulla facilisi. Aenean quis purus sagittis, dapibus nibh eget, ornare nunc. Donec posuere erat quis ipsum facilisis, quis porttitor dui cursus. Etiam convallis arcu sapien, vitae placerat diam molestie sit amet. Vivamus sapien augue, porta sed tortor ut, molestie ornare nisl. Nullam sed mi turpis. Donec sed finibus risus. Nunc interdum semper mauris quis vehicula. Phasellus in nisl faucibus, pellentesque massa vel, faucibus urna. Proin sed tortor lorem. Curabitur eu nisi semper, placerat tellus sed, varius nulla. Etiam luctus ac purus nec aliquet. Phasellus nisl metus, dictum ultricies justo a, laoreet consectetur risus. Vestibulum vulputate in felis ac blandit. Aliquam erat volutpat. Sed quis ultrices lectus. Curabitur at scelerisque elit, a bibendum nisi. Integer facilisis ex dolor, vel gravida metus vestibulum ac. Aliquam condimentum fermentum rhoncus. Nunc tortor arcu, condimentum non ex consequat, porttitor maximus est. Duis semper risus odio, quis feugiat sem elementum nec. Nam mattis nec dui sit amet volutpat. Sed facilisis, nunc quis porta consequat, ante mi tincidunt massa, eget euismod sapien nunc eget sem. Curabitur orci neque, eleifend at mattis quis, malesuada ac nibh. Vestibulum sed laoreet dolor, ac facilisis urna. Vestibulum luctus id nulla at auctor. Nunc pharetra massa orci, ut pharetra metus faucibus eget. Etiam eleifend, tellus id lobortis molestie, sem magna elementum dui, dapibus ullamcorper nisl enim ac urna. Nam posuere ullamcorper tellus, ac blandit nulla vestibulum nec. Vestibulum ornare, ligula quis aliquet cursus, metus nisi congue nulla, vitae posuere elit mauris at justo. Nullam ut fermentum arcu, nec laoreet ligula. Morbi quis consectetur nisl, nec consectetur justo. Curabitur eget eros hendrerit, ullamcorper dolor non, aliquam elit. Aliquam mollis justo vel aliquam interdum. Aenean bibendum rhoncus ante a commodo. Vestibulum bibendum sapien a accumsan pharetra... Curabitur condimentum consequat felis, id rutrum velit cursus vel. Proin nulla est, posuere in velit at, faucibus dignissim diam. Quisque quis erat euismod, malesuada erat eu, congue nisi. Ut risus lectus, auctor at libero sit amet, accumsan ultricies est. Donec eget iaculis enim. Nunc ac egestas tellus, nec efficitur magna. Sed nec nisl ut augue laoreet sollicitudin vitae nec quam. Vestibulum pretium nisl bibendum, tempor velit eu, semper velit. Nulla facilisi. Aenean quis purus sagittis, dapibus nibh eget, ornare nunc. Donec posuere erat quis ipsum facilisis, quis porttitor dui cursus. Etiam convallis arcu sapien, vitae placerat diam molestie sit amet. Vivamus sapien augue, porta sed tortor ut, molestie ornare nisl. Nullam sed mi turpis. Donec sed finibus risus. Nunc interdum semper mauris quis vehicula. Phasellus in nisl faucibus, pellentesque massa vel, faucibus urna. Proin sed tortor lorem. Curabitur eu nisi semper, placerat tellus sed, varius nulla. Etiam luctus ac purus nec aliquet. Phasellus nisl metus, dictum ultricies justo a, laoreet consectetur risus. Vestibulum vulputate in felis ac blandit. Aliquam erat volutpat. Sed quis ultrices lectus. Curabitur at scelerisque elit, a bibendum nisi. Integer facilisis ex dolor, vel gravida metus vestibulum ac. Aliquam condimentum fermentum rhoncus. Nunc tortor arcu, condimentum non ex consequat, porttitor maximus est. Duis semper risus odio, quis feugiat sem elementum nec. Nam mattis nec dui sit amet volutpat. Sed facilisis, nunc quis porta consequat, ante mi tincidunt massa, eget euismod sapien nunc eget sem. Curabitur orci neque, eleifend at mattis quis, malesuada ac nibh. Vestibulum sed laoreet dolor, ac facilisis urna. Vestibulum luctus id nulla at auctor. Nunc pharetra massa orci, ut pharetra metus faucibus eget. Etiam eleifend, tellus id lobortis molestie, sem magna elementum dui, dapibus ullamcorper nisl enim ac urna. Nam posuere ullamcorper tellus, ac blandit nulla vestibulum nec. Vestibulum ornare, ligula quis aliquet cursus, metus nisi congue nulla, vitae posuere elit mauris at justo. Nullam ut fermentum arcu, nec laoreet ligula. Morbi quis consectetur nisl, nec consectetur justo. Curabitur eget eros hendrerit, ullamcorper dolor non, aliquam elit. Aliquam mollis justo vel aliquam interdum. Aenean bibendum rhoncus ante a commodo. Vestibulum bibendum sapien a accumsan pharetra.';

	// Pre-Encoded (base64) images (if any)
	var checkGreen =
		'data:image/png;base64,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';
	var starlabsLogoSml =
		'data:image/png;base64,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';

	var svgBase64 =
		'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB2ZXJzaW9uPSIxLjEiIGlkPSJDYXBhXzEiIHg9IjBweCIgeT0iMHB4IiB3aWR0aD0iNTEycHgiIGhlaWdodD0iNTEycHgiIHZpZXdCb3g9IjAgMCA2NCA2NCIgc3R5bGU9ImVuYWJsZS1iYWNrZ3JvdW5kOm5ldyAwIDAgNjQgNjQ7IiB4bWw6c3BhY2U9InByZXNlcnZlIj48Zz48Zz48ZyBpZD0iY2lyY2xlX2NvcHlfNF8zXyI+PGc+PHBhdGggZD0iTTMyLDBDMTQuMzI3LDAsMCwxNC4zMjcsMCwzMmMwLDE3LjY3NCwxNC4zMjcsMzIsMzIsMzJzMzItMTQuMzI2LDMyLTMyQzY0LDE0LjMyNyw0OS42NzMsMCwzMiwweiBNMjguMjIyLDQxLjE5MSAgICAgIEwyOCw0MC45NzFsLTAuMjIyLDAuMjIzbC04Ljk3MS04Ljk3MWwxLjQxNC0xLjQxNUwyOCwzOC41ODZsMTUuNzc3LTE1Ljc3OGwxLjQxNCwxLjQxNEwyOC4yMjIsNDEuMTkxeiIgZmlsbD0iIzAwODhjYyIvPjwvZz48L2c+PC9nPjwvZz48L3N2Zz4=';
	var svgHyperlinkImage =
		'image/jpeg;base64,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';
}
var gArrNamesF = ['Markiplier','Jack','Brian','Paul','Ev','Ann','Michelle','Jenny','Lara','Kathryn'];
var gArrNamesL = ['Johnson','Septiceye','Lapston','Lewis','Clark','Griswold','Hart','Cube','Malloy','Capri'];
var gStrHello = 'BONJOUR - CIAO - GUTEN TAG - HELLO - HOLA - NAMASTE - OLÀ - ZDRAS-TVUY-TE - こんにちは - 你好';
var gOptsTabOpts = { x:0.4, y:0.13, w:12.5, colW:[9,3.5] };
var gOptsTextL = { color:'9F9F9F', margin:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] };
var gOptsOptsR = { color:'9F9F9F', margin:3, border:[0,0,{pt:'1',color:'CFCFCF'},0], align:'right' };
var gOptsTextR = { text:'PptxGenJS', options:gOptsOptsR };
var gOptsCode = { color:'9F9F9F', margin:3, border:{pt:'1',color:'CFCFCF'}, fill:'F1F1F1', fontFace:'Courier', fontSize:12 };
var gOptsSubTitle = { x:0.5, y:0.7, w:4, h:0.3, fontSize:18, fontFace:'Arial', color:'0088CC', fill:'FFFFFF' };
var gDemoTitleText = { fontSize:14, color:'0088CC', bold:true };
var gDemoTitleOpts = { fontSize:13, color:'9F9F9F' };
var gPaths = {
	'starlabsBkgd': { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.4.0/examples/images/starlabs_bkgd.jpg' },
	'starlabsLogo': { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.4.0/examples/images/starlabs_logo.png' },
	'wikimedia1'  : { path:'https://upload.wikimedia.org/wikipedia/en/a/a9/Example.jpg' },
	'wikimedia2'  : { path:'https://upload.wikimedia.org/wikipedia/commons/1/17/PNG-Gradient_hex.png' },
	'wikimedia_svg1': { path:'https://upload.wikimedia.org/wikipedia/commons/2/28/Cadenas-ferme-vert.svg' },
	'wikimedia_svg2': { path:'https://upload.wikimedia.org/wikipedia/commons/3/3b/Crystal_Clear_action_go.svg' },
	'ccCopyRemix'  : { path:'https://cdn.rawgit.com/gitbrent/PptxGenJS/v2.1.0/examples/images/cc_copyremix.gif' },
	'ccLogo'       : { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/images/cc_logo.jpg' },
	'ccLicenseComp': { path:'../SiteAssets/pptxgenjs/examples/images/cc_license_comp.png' },
	'ccDjGif'      : { path:'https://cdn.rawgit.com/gitbrent/PptxGenJS/master/examples/images/cc_dj.gif' },
	'gifAnimTrippy': { path:'https://cdn.rawgit.com/gitbrent/PptxGenJS/master/examples/images/trippy.gif' },
	'chicagoBean'  : { path:'https://upload.wikimedia.org/wikipedia/commons/thumb/d/d1/Chicago_Bean_Bohne_%2822038051679%29.jpg/256px-Chicago_Bean_Bohne_%2822038051679%29.jpg?op=paramTest&ampersandTest' },
	'sample_avi': { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/media/sample.avi' },
	'sample_m4v': { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/media/sample.m4v' },
	'sample_mov': { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/examples/media/sample.mov' },
	'sample_mp4': { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/media/sample.mp4' },
	'sample_mpg': { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/media/sample.mpg' },
	'sample_mp3': { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/media/sample.mp3' },
	'sample_wav': { path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/media/sample.wav' }
}

// ==================================================================================================================

function getTimestamp() {
	var dateNow = new Date();
	var dateMM = dateNow.getMonth() + 1; dateDD = dateNow.getDate(); dateYY = dateNow.getFullYear(), h = dateNow.getHours(); m = dateNow.getMinutes();
	return dateNow.getFullYear() +''+ (dateMM<=9 ? '0' + dateMM : dateMM) +''+ (dateDD<=9 ? '0' + dateDD : dateDD) + (h<=9 ? '0' + h : h) + (m<=9 ? '0' + m : m);
}

// ==================================================================================================================

function runEveryTest() {
	execGenSlidesFuncs( ['Master', 'Chart', 'Image', 'Media', 'Shape', 'Text', 'Table'] );

	// Dont run this automatically as Html2Pptx needs table to be visible as of 2.2.0
	// if ( typeof table2slides1 !== 'undefined' ) table2slides1();
}

function execGenSlidesFuncs(type) {
	// STEP 1: Instantiate new PptxGenJS object
	var pptx;
	if ( NODEJS ) {
		var PptxGenJsLib;
		var fs = require('fs');
		if (fs.existsSync('../dist/pptxgen.js')) {
			PptxGenJsLib = require('../dist/pptxgen.js'); // for LOCAL TESTING
		}
		else {
			PptxGenJsLib = require("pptxgenjs");
		}
		pptx = new PptxGenJsLib();
		var base64Images = require('./images/base64Images.js');
		LOGO_STARLABS = base64Images.LOGO_STARLABS();
	}
	else {
		pptx = new PptxGenJS();
	}
	//if (console.log) console.log(' * pptxgenjs ver: '+ pptx.version); // Loaded okay?

	// STEP 2: Set Presentation props (as QA test only - these are not required)
	pptx.setAuthor('Brent Ely');
	pptx.setCompany(CUST_NAME);
	pptx.setRevision('15');
	pptx.setSubject('PptxGenJS Test Suite Export');
	pptx.setTitle('PptxGenJS Test Suite Presentation');

	// STEP 3: Set layout
	pptx.setLayout('LAYOUT_WIDE');

	// STEP 4: Create Master Slides (from the old `pptxgen.masters.js` file - `gObjPptxMasters` items)
	{
		var objBkg = { path:(NODEJS ? gPaths.starlabsBkgd.path.replace(/http.+\/examples/, '../examples') : gPaths.starlabsBkgd.path) };
		var objImg = { path:(NODEJS ? gPaths.starlabsLogo.path.replace(/http.+\/examples/, '../examples') : gPaths.starlabsLogo.path), x:4.6, y:3.5, w:4, h:1.8 };

		// TITLE_SLIDE
		pptx.defineSlideMaster({
			title: 'TITLE_SLIDE',
			bkgd: objBkg,
			objects: [
				//{ 'line':  { x:3.5, y:1.0, w:6.0, h:0.0, line:'0088CC', lineSize:5 } },
				//{ 'chart': { type:'PIE', data:[{labels:['R','G','B'], values:[10,10,5]}], opts:{x:11.3, y:0.0, w:2, h:2, dataLabelFontSize:9} } },
				//{ 'image': { x:11.3, y:6.4, w:1.67, h:0.75, data:starlabsLogoSml } },
				{ 'rect':  { x: 0.0, y:5.7, w:'100%', h:0.75, fill:'F1F1F1' } },
				{ 'text':
					{ text:'Global IT & Services :: Status Report',
					options:{ x:0.0, y:5.7, w:'100%', h:0.75, fontFace:'Arial', color:'363636', fontSize:20, align:'c', valign:'m', margin:0 } }
				}
			]
		});

		// MASTER_PLAIN
		pptx.defineSlideMaster({
			title: 'MASTER_PLAIN',
			bkgd: 'FFFFFF',
			margin:  [ 0.5, 0.25, 1.0, 0.25 ],
			objects: [
				{ 'rect':  { x: 0.00, y:6.90, w:'100%', h:0.6, fill:'003b75' } },
				{ 'image': { x:11.45, y:5.95, w:1.67, h:0.75, data:starlabsLogoSml } },
				{ 'text':
					{
						options: {x:0, y:6.9, w:'100%', h:0.6, align:'c', valign:'m', color:'FFFFFF', fontSize:12},
						text: 'S.T.A.R. Laboratories - Confidential'
					}
				}
			],
			slideNumber: { x:0.6, y:7.1, color:'FFFFFF', fontFace:'Arial', fontSize:10 }
		});

		// MASTER_SLIDE (MASTER_PLACEHOLDER)
		pptx.defineSlideMaster({
			title: 'MASTER_SLIDE',
			bkgd: 'FFFFFF',
			margin:  [ 0.5, 0.25, 1.0, 0.25 ],
			slideNumber: { x:0.6, y:7.1, color:'FFFFFF', fontFace:'Arial', fontSize:10 },
			objects: [
				{ 'rect':  { x: 0.00, y:6.90, w:'100%', h:0.6, fill:'003b75' } },
				{ 'image': { x:11.45, y:5.95, w:1.67, h:0.75, data:starlabsLogoSml } },
				{ 'text':
					{
						options: {x:0, y:6.9, w:'100%', h:0.6, align:'c', valign:'m', color:'FFFFFF', fontSize:12},
						text: 'S.T.A.R. Laboratories - Confidential'
					}
				},
				{ 'placeholder':
					{
						options: { name:'title', type:'title', x:0.6, y:0.2, w:12, h:1.0 },
						text: ''
					}
				},
				{ 'placeholder':
					{
						options: { name:'body', type:'body', x:0.6, y:1.5, w:12, h:5.25 },
						text: '(supports custom placeholder text!)'
					}
				}
			]
		});

		// THANKS_SLIDE (THANKS_PLACEHOLDER)
		pptx.defineSlideMaster({
			title: 'THANKS_SLIDE',
			bkgd: '36ABFF',
			objects: [
				{ 'rect':  { x:0.0, y:3.4, w:'100%', h:2.0, fill:'ffffff' } },
				{ 'placeholder': { options:{ name:'thanksText', type:'title', x:0.0, y:0.9, w:'100%', h:1, fontFace:'Arial', color:'FFFFFF', fontSize:60, align:'c' } } },
				{ 'image': objImg },
				{ 'placeholder':
					{
						options: { name:'body', type:'body', x:0.0, y:6.45, w:'100%', h:1, fontFace:'Courier', color:'FFFFFF', fontSize:32, align:'c' },
						text: '(add homepage URL)'
					}
				}
			]
		});

		// MISC: Only used for Issues, ad-hoc slides etc (for screencaps)
		pptx.defineSlideMaster({
			title: 'DEMO_SLIDE',
			objects: [
				{ 'rect':  { x:0.0, y:7.1, w:'100%', h:0.4, fill:'f1f1f1' } },
				{ 'text':  { text:'PptxGenJS - JavaScript PowerPoint Library - (github.com/gitbrent/PptxGenJS)', options:{ x:0.0, y:7.1, w:'100%', h:0.4, color:'6c6c6c', fontSize:10, align:'c' } } }
			]
		});
	}

	// STEP 5: Run requested test
	var arrTypes = ( typeof type === 'string' ? [type] : type );
	arrTypes.forEach(function(type,idx){ eval( 'genSlides_'+type+'(pptx)' ); });

	// LAST: Export Presentation
	if ( NODEJS ) {
		pptx.save('PptxGenJS_Demo_Node_'+type+'_'+getTimestamp());
		console.log('\nDemo file created:\n * '+'PptxGenJS_Demo_Node_'+type+'_'+getTimestamp());
	}
	else {
		pptx.save('PptxGenJS_Demo_Browser_'+type+'_'+getTimestamp());
	}
}

// ==================================================================================================================

function genSlides_Table(pptx) {
	// SLIDE 1: Table text alignment and cell styles
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs:\nhttps://gitbrent.github.io/PptxGenJS/docs/api-tables.html');
		slide.addTable( [ [{ text:'Table Examples 1', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// DEMO: align/valign -------------------------------------------------------------------------
		var objOpts1 = { x:0.5, y:0.7, w:4, fontSize:18, fontFace:'Arial', color:'0088CC' };
		slide.addText('Cell Text Alignment:', objOpts1);

		var arrTabRows = [
			[
				{ text: 'Top Lft', options: { valign:'top',    align:'left'  , fontFace:'Arial'   } },
				{ text: 'Top Ctr', options: { valign:'t'  ,    align:'center', fontFace:'Courier' } },
				{ text: 'Top Rgt', options: { valign:'t'  ,    align:'right' , fontFace:'Verdana' } }
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
			arrTabRows, { x:0.5, y:1.1, w:5.0, rowH:0.75, fill:'F7F7F7', fontSize:14, color:'363636', border:{pt:'1', color:'BBCCDD'} }
		);
		// Pass default cell style as tabOpts, then just style/override individual cells as needed

		// DEMO: cell styles --------------------------------------------------------------------------
		var objOpts2 = { x:6.0, y:0.7, w:4, fontSize:18, fontFace:'Arial', color:'0088CC' };
		slide.addText('Cell Styles:', objOpts2);

		var arrTabRows = [
			[
				{ text: 'White',  options: { fill:'6699CC', color:'FFFFFF' } },
				{ text: 'Yellow', options: { fill:'99AACC', color:'FFFFAA' } },
				{ text: 'Pink',   options: { fill:'AACCFF', color:'E140FE' } }
			],
			[
				{ text: '12pt', options: { fill:'FF0000', fontSize:12 } },
				{ text: '20pt', options: { fill:'00FF00', fontSize:20 } },
				{ text: '28pt', options: { fill:'0000FF', fontSize:28 } }
			],
			[
				{ text: 'Bold',      options: { fill:'003366', bold:true } },
				{ text: 'Underline', options: { fill:'336699', underline:true } },
				{ text: '10pt Pad',  options: { fill:'6699CC', margin:10 } }
			]
		];
		slide.addTable(
			arrTabRows, { x:6.0, y:1.1, w:7.0, rowH:0.75, fill:'F7F7F7', color:'FFFFFF', fontSize:16, valign:'center', align:'ctr', border:{pt:'1', color:'FFFFFF'} }
		);

		// DEMO: Row/Col Width/Heights ----------------------------------------------------------------
		var objOpts3 = { x:0.5, y:3.6, fontSize:18, fontFace:'Arial', color:'0088CC' };
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
				fontSize:14, valign:'center', align:'ctr',
				border:{pt:'1', color:'BBCCDD'}
			}
		);
	}

	// SLIDE 2: Table row/col-spans
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html');
		// 2: Slide title
		slide.addTable([ [{ text:'Table Examples 2', options:gOptsTextL },gOptsTextR] ], { x:'4%', y:'2%', w:'95%', h:'4%' }); // QA: this table's x,y,w,h all using %

		// DEMO: Rowspans/Colspans ----------------------------------------------------------------
		var optsSub = JSON.parse(JSON.stringify(gOptsSubTitle));
		slide.addText('Colspans/Rowspans:', optsSub);

		var tabOpts1 = { x:0.5, y:1.1, w:'90%', h:2, fill:'F5F5F5', color:'3D3D3D', fontSize:16, border:{pt:4, color:'FFFFFF'}, align:'c', valign:'m' };
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

		var tabOpts2 = { x:0.5, y:3.3, w:12.4, h:1.5, fontSize:14, fontFace:'Courier', align:'center', valign:'middle', fill:'F9F9F9', border:{pt:'1',color:'c7c7c7'}};
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

		var tabOpts3 = {x:0.5, y:5.15, w:6.25, h:2, margin:0.25, align:'center', valign:'middle', fontSize:16, border:{pt:'1',color:'c7c7c7'}, fill:'F1F1F1' }
		var arrTabRows3 = [
			[ {text:'A1\nA2\nA3', options:{rowspan:3, fill:'FFFCCC'}}, {text:'B1\nB2', options:{rowspan:2, fill:'FFFCCC'}}, 'C1' ],
			[ 'C2' ],
			[ { text:'B3 -> C3', options:{colspan:2, fill:'99FFCC'} } ]
		];
		slide.addTable(arrTabRows3, tabOpts3);

		var tabOpts4 = {x:7.4, y:5.15, w:5.5, h:2, margin:0, align:'center', valign:'middle', fontSize:16, border:{pt:'1',color:'c7c7c7'}, fill:'F2F9FC' }
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
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html');
		slide.addTable( [ [{ text:'Table Examples 3', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// DEMO: Rowspans/Colspans ----------------------------------------------------------------
		var optsSub = JSON.parse(JSON.stringify(gOptsSubTitle));
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

		var taboptions5 = { x:0.6, y:1.3, w:'90%', h:5.5, margin:0, fontSize:14, align:'c', valign:'m', border:{pt:'1'} };

		slide.addTable(arrTabRows5, taboptions5);
	}

	// SLIDE 4: Cell Formatting / Cell Margins
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html');
		// 2: Slide title
		slide.addTable( [ [{ text:'Table Examples 4', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// Cell Margins
		var optsSub = JSON.parse(JSON.stringify(gOptsSubTitle));
		slide.addText('Cell Margins:', optsSub);

		slide.addTable( [['margin:0']],           { x:0.5, y:1.1, margin:0,           w:1.2, fill:'FFFCCC' } );
		slide.addTable( [['margin:[0,0,0,20]']],  { x:2.5, y:1.1, margin:[0,0,0,20],  w:2.0, fill:'FFFCCC', align:'r' } );
		slide.addTable( [['margin:5']],           { x:5.5, y:1.1, margin:5,           w:1.0, fill:'F1F1F1' } );
		slide.addTable( [['margin:[40,5,5,20]']], { x:7.5, y:1.1, margin:[40,5,5,20], w:2.2, fill:'F1F1F1' } );
		slide.addTable( [['margin:[80,5,5,10]']], { x:10.5,y:1.1, margin:[80,5,5,10], w:2.2, fill:'F1F1F1' } );

		slide.addTable( [{text:'number zero:', options:{margin:5}}, {text:0, options:{margin:5}}], { x:0.5, y:1.9, w:3, fill:'f2f9fc', border:'none', colW:[2,1] } );
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
		var optsSub = JSON.parse(JSON.stringify(gOptsSubTitle)); optsSub.y = 3.9;
		slide.addText('Complex Cell Border:', optsSub);
		var arrBorder = [ {color:'FF0000',pt:1}, {color:'00ff00',pt:3}, {color:'0000ff',pt:5}, {color:'9e9e9e',pt:7} ];
		slide.addTable( [['Borders!']], { x:0.5, y:4.3, w:12.3, rowH:1.5, fill:'F5F5F5', color:'3D3D3D', fontSize:18, border:arrBorder, align:'c', valign:'c' } );

		// Invalid char check
		var optsSub = JSON.parse(JSON.stringify(gOptsSubTitle)); optsSub.y = 6.1;
		slide.addText('Escaped Invalid Chars:', optsSub);
		var arrTabRows3 = [['<', '>', '"', "'", '&', 'plain']];
		slide.addTable( arrTabRows3, { x:0.5, y:6.5, w:12.3, rowH:0.5, fill:'F5F5F5', color:'3D3D3D', border:'FFFFFF', align:'c', valign:'c' } );

	}

	// SLIDE 5: Cell Word-Level Formatting
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html');
		slide.addTable( [ [{ text:'Table Examples 5', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );
		slide.addText(
			'The following textbox and table cell use the same array of text/options objects, making word-level formatting familiar and consistent across the library.',
			{ x:0.5, y:0.5, w:'95%', h:0.5, margin:0.1, fontSize:14 }
		);
		slide.addText("[\n"
			+ "  { text:'1st line', options:{ fontSize:24, color:'99ABCC', align:'r', breakLine:true } },\n"
			+ "  { text:'2nd line', options:{ fontSize:36, color:'FFFF00', align:'c', breakLine:true } },\n"
			+ "  { text:'3rd line', options:{ fontSize:48, color:'0088CC', align:'l' } }\n"
			+ "]",
			{ x:1, y:1.1, w:11, h:1.5, margin:0.1, fontFace:'Courier', fontSize:14, fill:'F1F1F1', color:'333333' }
		);

		// Textbox: Text word-level formatting
		slide.addText('Textbox:', { x:1, y:2.8, w:3, fontSize:18, fontFace:'Arial', color:'0088CC' });

		var arrTextObjects = [
			{ text:'1st line', options:{ fontSize:24, color:'99ABCC', align:'r', breakLine:true } },
			{ text:'2nd line', options:{ fontSize:36, color:'FFFF00', align:'c', breakLine:true } },
			{ text:'3rd line', options:{ fontSize:48, color:'0088CC', align:'l' } }
		];
		slide.addText( arrTextObjects, { x:2.5, y:2.8, w:9, h:2, margin:0.1, fill:'232323' } );

		// Table cell: Use the exact same code from addText to do the same word-level formatting within a cell
		slide.addText('Table:', { x:1, y:5, w:3, fontSize:18, fontFace:'Arial', color:'0088CC' });

		var opts2 = { x:2.5, y:5, w:9, h:2, align:'center', valign:'middle', colW:[1.5,1.5,6], border:{pt:'1'}, fill:'F1F1F1' }
		var arrTabRows = [
			[
				{ text:'Cell 1A',       options:{fontFace:'Arial'  } },
				{ text:'Cell 1B',       options:{fontFace:'Courier'} },
				{ text: arrTextObjects, options:{fill:'232323'      } }
			]
		];
		slide.addTable(arrTabRows, opts2);
	}

	// SLIDE 6: Cell Word-Level Formatting
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html');
		slide.addTable( [ [{ text:'Table Examples 6', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		var optsSub = JSON.parse(JSON.stringify(gOptsSubTitle));
		slide.addText('Table Cell Word-Level Formatting:', optsSub);

		// EX 1:
		var arrCell1 = [{ text:'Cell 1A', options:{ color:'0088cc' } }];
		var arrCell2 = [{ text:'Red ', options:{color:'FF0000'} }, { text:'Green ', options:{color:'00FF00'} }, { text:'Blue', options:{color:'0000FF'} }];
		var arrCell3 = [{ text:'Bullets\nBullets\nBullets', options:{ color:'0088cc', bullet:true } }];
		var arrCell4 = [{ text:'Numbers\nNumbers\nNumbers', options:{ color:'0088cc', bullet:{type:'number'} } }];
		var arrTabRows = [
			[{ text:arrCell1 }, { text:arrCell2, options:{valign:'m'} }, { text:arrCell3, options:{valign:'m'} }, { text:arrCell4, options:{valign:'b'} }]
		];
		slide.addTable( arrTabRows, { x:0.6, y:1.25, w:12, h:3, fontSize:24, border:{pt:'1'}, fill:'F1F1F1' } );

		// EX 2:
		slide.addTable(
			[
				{ text:[
					{ text:'I am a text object with bullets ', options:{color:'CC0000', bullet:{code:'2605'}} },
					{ text:'and i am the next text object'   , options:{color:'00CD00', bullet:{code:'25BA'}} },
					{ text:'Final text object w/ bullet:true', options:{color:'0000AB', bullet:true} }
				]},
				{ text:[
					{ text:'Cell', options:{fontSize:36, align:'l', breakLine:true} },
					{ text:'#2',   options:{fontSize:60, align:'r', color:'CD0101'} }
				]},
				{ text:[
					{ text:'Cell', options:{fontSize:36, fontFace:'Courier', color:'dd0000', breakLine:true} },
					{ text:'#'   , options:{fontSize:60, color:'8648cd'} },
					{ text:'3'   , options:{fontSize:60, color:'33ccef'} }
				]}
			],
			{ x:0.6, y:4.75, w:12, h:2, fontSize:24, colW:[8,2,2], valign:'m', border:{pt:'1'}, fill:'F1F1F1' }
		);
	}

	// SLIDE 7+: Table auto-paging
	// ======== -----------------------------------------------------------------------------------
	{
		var arrRows = [];
		var arrText = [];
		for (var idx=0; idx<gArrNamesF.length; idx++) {
			var strText = ( idx == 0 ? gStrLoremIpsum.substring(0,100) : gStrLoremIpsum.substring(idx*100,idx*200) );
			arrRows.push( [idx, gArrNamesF[idx], strText] );
			arrText.push( [strText] );
		}

		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-tables.html');
		slide.addText( [{text:'Table Examples: ', options:gDemoTitleText},{text:'Auto-Paging Example', options:gDemoTitleOpts}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:0.5, y:0.6, colW:[0.75,1.75,10], margin:2, border:'CFCFCF' } );

		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:gDemoTitleText},{text:'Smaller Table Area', options:gDemoTitleOpts}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:3.0, y:0.6, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF' } );

		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:gDemoTitleText},{text:'Test: Correct starting Y location upon paging', options:gDemoTitleOpts}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:3.0, y:4.0, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF' } );

		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:gDemoTitleText},{text:'Test: `{ newPageStartY: 1.5 }`', options:gDemoTitleOpts}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:3.0, y:4.0, newPageStartY:1.5, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF' } );

		var slide = pptx.addNewSlide('MASTER_PLAIN', {bkgd:'CCFFCC'});
		slide.addText( [{text:'Table Examples: ', options:gDemoTitleText},{text:'Master Page with Auto-Paging', options:gDemoTitleOpts}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:1.0, y:0.6, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF' } );

		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:gDemoTitleText},{text:'Auto-Paging Disabled', options:gDemoTitleOpts}], {x:0.5, y:0.13, w:'90%'} );
		slide.addTable( arrRows, { x:1.0, y:0.6, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF', autoPage:false } );

		// lineWeight option demos
		var slide = pptx.addNewSlide();
		slide.addText( [{text:'Table Examples: ', options:gDemoTitleText},{text:'lineWeight:0', options:gDemoTitleOpts}], {x:0.5, y:0.13, w:3} );
		slide.addTable( arrText, { x:0.50, y:0.6, w:4, margin:5, border:'CFCFCF', autoPage:true } );

		slide.addText( [{text:'Table Examples: ', options:gDemoTitleText},{text:'lineWeight:0.5', options:gDemoTitleOpts}], {x:5.0, y:0.13, w:3} );
		slide.addTable( arrText, { x:4.75, y:0.6, w:4, margin:5, border:'CFCFCF', autoPage:true, lineWeight:0.5 } );

		slide.addText( [{text:'Table Examples: ', options:gDemoTitleText},{text:'lineWeight:-0.5', options:gDemoTitleOpts}], {x:9.0, y:0.13, w:3} );
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
	// Create a gap for testing `displayBlanksAs` in line charts (2.3.0)
	arrDataLineStat[2].values = [55, null, null, 55];

	// SLIDE 1: Bar Chart ------------------------------------------------------------------
	function slide1() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Bar Chart', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

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
			// valAxisCrossesAt: 10,

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

			showLegend : true,
			legendPos  :  't',
			legendColor: 'FF0000',
			showTitle  : true,
			titleColor : 'FF0000',
			title      : 'Red Title and Legend'
		};
		slide.addChart( pptx.charts.BAR, arrDataHighVals, optsChartBar4 );
	}

	// SLIDE 2: Bar Chart Grid/Axis Options ------------------------------------------------
	function slide2() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Bar Chart Grid/Axis Options', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

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

			valAxisOrientation: 'maxMin',
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

			catAxisLabelColor    : '0000CC',
			catAxisLabelFontFace : 'Times',
			catAxisLabelFontSize : 11,
			catAxisLabelFrequency: 1,
			catAxisOrientation   : 'minMax',

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
	function slide3() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Bar Chart: Stacked/PercentStacked and Data Table', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

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

			catAxisOrientation  : 'maxMin',
			catAxisLabelColor   : 'CC0000',
			catAxisLabelFontFace: 'Helvetica Neue',
			catAxisLabelFontSize: 14,
			catAxisLabelFontBold: true,
			valAxisLabelFontBold: true,

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
			dataLabelFontBold: true,
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
			barDir     : 'bar',
			barGrouping: 'percentStacked',
			dataBorder   : { pt:'1', color:'F1F1F1' },
			catAxisHidden: true,
			valAxisHidden: true,
			showTitle    : false,
			layout       : {x:0.1, y:0.1, w:1, h:1},
			showDataTable:           true,
			showDataTableKeys:       true,
			showDataTableHorzBorder: false,
			showDataTableVertBorder: false,
			showDataTableOutline:    false
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
			legendPos :  't',

			showDataTable:     true,
			showDataTableKeys: false
		};
		slide.addChart( pptx.charts.BAR, arrDataHighVals, optsChartBar4 );
	}

	// SLIDE 4: Bar Chart - Lots of Bars ---------------------------------------------------
	function slide4() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Lots of Bars (>26 letters)', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		var arrDataHighVals = [
			{
				name  : 'TEST: getExcelColName',
				labels: LETTERS.concat(['AA','AB','AC','AD']),
				values: [-5,-3,0,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30 ]
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

			chartColors: ['EE1122'],
			invertedColors: ['0088CC'],

			showValAxisTitle: true,
			valAxisTitle: "Column Index",
			valAxisTitleColor: "c11c13",
			valAxisTitleFontSize: 16,
		};

		// TEST `getExcelColName()` to ensure Excel Column names are generated correctly above >26 chars/cols
		slide.addChart(pptx.charts.BAR, arrDataHighVals, optsChart);
	}

	// SLIDE 5: Bar Chart: Data Series Colors, majorUnits, and valAxisLabelFormatCode ------
	function slide5() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Multi-Color Bars, `catLabelFormatCode`, `valAxisMajorUnit`, `valAxisLabelFormatCode`', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// TOP-LEFT
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name  : 'Labels are Excel Date Values',
					labels: [37987,38018,38047,38078,38108,38139],
					values: [20, 30, 10, 25, 15, 5]
				}
			],
			{
				x:0.5, y:0.6, w:'45%', h:3,
				barDir: 'bar',
				chartColors: ['0077BF','4E9D2D','ECAA00','5FC4E3','DE4216','154384'],
				catLabelFormatCode: "yyyy-mm",
				valAxisMajorUnit: 15,
				valAxisMaxVal: 45,
				showTitle: true,
				titleFontSize: 14,
				titleColor: '0088CC',
				title: 'Bar Charts Can Be Multi-Color'
			}
		);

		// TOP-RIGHT
		// NOTE: Labels are ppt/excel dates (days past 1900)
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name  : 'Too Many Colors Series',
					labels: [37987,38018,38047,38078,38108,38139],
					values: [.20, .30, .10, .25, .15, .05]
				}
			],
			{
				x:7, y:0.6, w:'45%', h:3,
				valAxisMaxVal:1,
				barDir: 'bar',
				catAxisLineShow: false,
				valAxisLineShow: false,
				showValue: true,
				catLabelFormatCode: "mmm-yy",
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
			{  x:0.5, y:4.0, w:'45%', h:3,
				barDir: 'bar',
				showValue: true,
				dataLabelPosition: 'outEnd',
				dataLabelFormatCode: '#%',
				valAxisLabelFormatCode: '0.#0',
				chartColors: ['0077BF','4E9D2D','ECAA00'],
				valAxisMaxVal: .40,
				barGapWidthPct: 50
			}
		);

		// BOTTOM-RIGHT
		slide.addChart(
			pptx.charts.BAR,
			[
				{
					name  : 'Escaped XML Chars',
					labels: ['Es', 'cap', 'ed', 'XML', 'Chars', "'", '"', '&', '<', '>'],
					values: [1.20, 2.30, 3.10, 4.25, 2.15, 6.05, 8.01, 2.02, 9.9, 0.9]
				}
			],
			{
				x:7, y:4, w:'45%', h:3,
				barDir: 'bar',
				showValue: true,
				dataLabelPosition: 'outEnd',
				chartColors: ['0077BF','4E9D2D','ECAA00','5FC4E3','DE4216','154384','7D666A','A3C961','EF907B','9BA0A3'],
				barGapWidthPct: 25,
				catAxisOrientation: 'maxMin',
				valAxisOrientation: 'maxMin',
				valAxisMaxVal: 10,
				valAxisMajorUnit: 1
			}
		);
	}

    // SLIDE 6: 3D Bar Chart ---------------------------------------------------------------
    function slide6() {
        var slide = pptx.addNewSlide();
        slide.addTable( [ [{ text:'Chart Examples: 3D Bar Chart', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

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
            catAxisLabelFontFace: 'Arial',
            catAxisLabelFontSize: 10,
            catAxisOrientation  : 'maxMin',

            serAxisLabelColor   : '00EE00',
            serAxisLabelFontFace: 'Arial',
            serAxisLabelFontSize: 10
        };
        slide.addChart( pptx.charts.BAR3D, arrDataRegions, optsChartBar1 );

        // TOP-RIGHT: V/col
        var optsChartBar2 = { x:7.0, y:0.6, w:6.0, h:3.0,
            barDir: 'col',
            bar3DShape: 'cylinder',
            catAxisLabelColor   : '0000CC',
            catAxisLabelFontFace: 'Courier',
            catAxisLabelFontSize: 12,

            dataLabelColor     : '000000',
            dataLabelFontFace  : 'Arial',
            dataLabelFontSize  : 11,
            dataLabelPosition  : 'outEnd',
            dataLabelFormatCode: '#.0',
            dataLabelBkgrdColors: true,
            showValue          : true
        };
        slide.addChart( pptx.charts.BAR3D, arrDataRegions, optsChartBar2 );

        // BTM-LEFT: H/bar - TITLE and LEGEND
        slide.addText( '.', { x:0.5, y:3.8, w:6.0, h:3.5, fill:'F1F1F1', color:'F1F1F1'} );
        var optsChartBar3 = { x:0.5, y:3.8, w:6.0, h:3.5,
            barDir: 'col',
            bar3DShape: 'pyramid',
            barGrouping: 'stacked',

            catAxisLabelColor   : 'CC0000',
            catAxisLabelFontFace: 'Arial',
            catAxisLabelFontSize: 10,

            showValue          : true,
            dataLabelBkgrdColors: true,

            showTitle : true,
            title: 'Sales by Region'
        };
        slide.addChart( pptx.charts.BAR3D, arrDataHighVals, optsChartBar3 );

        // BTM-RIGHT: V/col - TITLE and LEGEND
        slide.addText( '.', { x:7.0, y:3.8, w:6.0, h:3.5, fill:'F1F1F1', color:'F1F1F1'} );
        var optsChartBar4 = { x:7.0, y:3.8, w:6.0, h:3.5,
            barDir: 'col',
            bar3DShape: 'coneToMax',
            chartColors: ['0088CC', '99FFCC'],

            catAxisLabelColor   : '0000CC',
            catAxisLabelFontFace: 'Times',
            catAxisLabelFontSize: 11,
            catAxisOrientation  : 'minMax',

            dataBorder         : { pt:'1', color:'F1F1F1' },
            dataLabelColor     : '000000',
            dataLabelFontFace  : 'Arial',
            dataLabelFontSize  : 10,
            dataLabelPosition  : 'ctr',

            showLegend : true,
            legendPos  :  't',
            legendColor: 'FF0000',
            showTitle  : true,
            titleColor : 'FF0000',
            title      : 'Red Title and Legend'
        };
        slide.addChart( pptx.charts.BAR3D, arrDataHighVals, optsChartBar4 );
    }

    // SLIDE 7: Tornado Chart --------------------------------------------------------------
	function slide7() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Tornado Chart - Grid and Axis Formatting', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

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
				valAxisMaxVal: 1,
				barDir: 'bar',
				axisLabelFormatCode: '#%',
				gridLineColor: 'D8D8D8',
				axisLineColor: 'D8D8D8',
				catAxisLineShow: false,
				valAxisLineShow: false,
				barGrouping: 'stacked',
				catAxisLabelPos: 'low',
				valueBarColors: true,
				shadow: 'none',
				chartColors: ['0077BF','4E9D2D','ECAA00','5FC4E3','DE4216','154384','7D666A','A3C961','EF907B','9BA0A3'],
				invertedColors: ['0065A2','428526','C99100','51A7C1','BD3813','123970','6A575A','8BAB52','CB7A69','84888B'],
				barGapWidthPct: 25,
				valAxisMajorUnit: 0.2
			}
		);
	}

	// SLIDE 8: Line Chart: Line Smoothing, Line Size, Symbol Size -------------------------
	function slide8() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Line Smoothing, Line Size, Line Shadow, Symbol Size', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

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
			shadow: 'none',
			//displayBlanksAs: 'gap', //uncomment only for test - looks broken otherwise!
			showLegend: true, legendPos: 'l'
		};
		slide.addChart( pptx.charts.LINE, arrDataLineStat, optsChartLine1 );

		// QA: DEMO: Test shadow option
		var shadowOpts = { type:'outer', color:'cd0011', blur:3, offset:12, angle:75, opacity:0.8 };
		var optsChartLine2 = { x:7.0, y:4.0, w:6.0, h:3.0,
			chartColors: [ COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK ],
			lineDataSymbolSize: 20,
			shadow: shadowOpts,
			showLegend: true, legendPos: 'b'
		};
		slide.addChart( pptx.charts.LINE, arrDataLineStat, optsChartLine2 );
	}

	// SLIDE 9: Line Chart: TEST: `lineDataSymbol` + `lineDataSymbolSize` ------------------
	function slide9() {
		var intWgap = 4.25;
		var opts_lineDataSymbol = ['circle','dash','diamond','dot','none','square','triangle'];
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Line Chart: lineDataSymbol option test', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

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

	// SLIDE 10: Line Chart: Lots of Cats --------------------------------------------------
	function slide10() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Line Chart: Lots of Lines', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

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

	// SLIDE 11: Area Chart: Misc ----------------------------------------------------------
	function slide11() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Area Chart', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

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
		var optsChartLine1 = {
			x:0.5, y:0.6, w:'45%', h:3,
			catAxisLabelRotate: 45,
			catAxisOrientation:'maxMin', valAxisOrientation:'maxMin'
		};
		slide.addChart( pptx.charts.AREA, arrDataTimeline2ser, optsChartLine1 );

		// TOP-RIGHT
		var optsChartLine2 = { x:7, y:0.6, w:'45%', h:3,
			chartColors: ['0088CC', '99FFCC'],
			chartColorsOpacity: 25,
			valAxisLabelRotate: 5,
			dataBorder: {pt:2, color:'FFFFFF'},
			fill: 'D1E1F1'
		};
		slide.addChart( pptx.charts.AREA, arrDataTimeline2ser, optsChartLine2 );

		// BOTTOM-LEFT
		var optsChartLine3 = { x:0.5, y:4.0, w:'45%', h:3,
			chartColors: ['0088CC', '99FFCC'],
			chartColorsOpacity: 50,
			valAxisLabelFormatCode: '#,K'
		};
		slide.addChart( pptx.charts.AREA, arrDataTimeline2ser, optsChartLine3 );

		// BOTTOM-RIGHT
		var optsChartLine4 = { x:7, y:4.0, w:'45%', h:3,
			chartColors: ['CC8833', 'CCFF69'],
			chartColorsOpacity: 75
		};
		slide.addChart( pptx.charts.AREA, arrDataTimeline2ser, optsChartLine4 );
	}

	// SLIDE 12: Pie Charts: All 4 Legend Options ------------------------------------------
	function slide12() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Pie Charts: Legends', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// INTERNAL USE: Not visible to user (its behind a chart): Used for ensuring ref counting works across obj types (eg: `rId` check/test)
		if (NODEJS) slide.addImage({ path:(NODEJS ? gPaths.ccCopyRemix.path.replace(/http.+\/examples/, '../examples') : gPaths.ccCopyRemix.path), x:0.5, y:1.0, w:1.2, h:1.2 });

		// TOP-LEFT
		slide.addText( '.', {x:0.5, y:0.5, w:4.2, h:3.2, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart(
			pptx.charts.PIE, dataChartPieStat,
			{x:0.5, y:0.5, w:4.2, h:3.2, showLegend:true, legendPos:'l', legendFontFace:'Courier New'}
		);

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
		// DEMO: `legendFontSize`, `titleAlign`, `titlePos`
		slide.addChart( pptx.charts.PIE, dataChartPieLocs,
		{
			x:9.8, y:0.5, w:3.2, h:3.2, dataBorder:{pt:'1',color:'F1F1F1'},
			showLegend: true,
			legendPos: 't',
			showTitle: true,
			title:'Left Title & Large Legend',

			legendFontSize: 14,
			titleAlign: 'left',
			titlePos: {x: 0, y: 0}
		});

		// BOTH: BTM-RIGHT
		slide.addText( '.', {x:9.8, y:4.0, w:3.2, h:3.2, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart( pptx.charts.PIE, dataChartPieLocs, {x:9.8, y:4.0, w:3.2, h:3.2, dataBorder:{pt:'1',color:'F1F1F1'}, showLegend:true, legendPos:'b', showTitle:true, title:'Title & Legend'} );
	}

	// SLIDE 13: Doughnut Chart ------------------------------------------------------------
	function slide13() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Doughnut Chart', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		var optsChartPie1 = { x:0.5, y:1.0, w:6.0, h:6.0,
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
			title: 'Resource Totals by Location',
			shadow: {
				offset: 20,
				blur: 20,
				type: 'inner'
			}
		};
		slide.addChart(pptx.charts.DOUGHNUT, dataChartPieLocs, optsChartPie2 );
	}

	// SLIDE 14: XY Scatter Chart ----------------------------------------------------------
	function slide14() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: XY Scatter Chart', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		var arrDataScatter1 = [
			{ name:'X-Axis',    values:[0,1,2,3,4,5,6,7,8,9] },
			{ name:'Y-Value 1', values:[13, 20, 21, 25] },
			{ name:'Y-Value 2', values:[21, 22, 25, 49] }
		];
		var arrDataScatter2 = [
			{ name:'X-Axis',   values:[1, 2, 3, 4, 5, 6] },
			{ name:'Airplane', values:[33, 20, 51, 65, 71, 75] },
			{ name:'Train',    values:[99, 88, 77, 89, 99, 99] },
			{ name:'Bus',      values:[21, 22, 25, 49, 59, 69] }
		];
		var arrDataScatterLabels = [
		    { name:'X-Axis',    values:[1, 10, 20, 30, 40, 50] },
		    { name:'Y-Value 1', values:[11, 23, 31, 45], labels:['Red 1', 'Red 2', 'Red 3', 'Red 4'] },
		    { name:'Y-Value 2', values:[21, 38, 47, 59], labels:['Blue 1', 'Blue 2', 'Blue 3', 'Blue 4'] }
		];

		// TOP-LEFT
		var optsChartScat1 = { x:0.5, y:0.6, w:'45%', h:3,
			valAxisTitle        : "Renters",
			valAxisTitleColor   : "428442",
			valAxisTitleFontSize: 14,
			showValAxisTitle    : true,
			lineSize: 0,
			catAxisTitle        : "Last 10 Months",
			catAxisTitleColor   : "428442",
			catAxisTitleFontSize: 14,
			showCatAxisTitle    : true
		};
		slide.addChart( pptx.charts.SCATTER, arrDataScatter1, optsChartScat1 );

		// TOP-RIGHT
		var optsChartScat2 = { x:7.0, y:0.6, w:'45%', h:3,
			fill: 'f1f1f1',
			showLegend: true,
			legendPos : 'b',

			lineSize  : 8,
			lineSmooth: true,
			lineDataSymbolSize: 12,
			lineDataSymbolLineColor: 'FFFFFF',

			chartColors: [ COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK ],
			chartColorsOpacity: 25
		};
		slide.addChart( pptx.charts.SCATTER, arrDataScatter2, optsChartScat2 );

		// BOTTOM-LEFT: (Labels)
		var optsChartScat3 = { x:0.5, y:4.0, w:'45%', h:3,
			fill: 'f2f9fc',
			//catAxisOrientation: 'maxMin',
			//valAxisOrientation: 'maxMin',
			showCatAxisTitle: false,
			showValAxisTitle: false,
			lineSize: 0,

			catAxisTitle        : "Data Point Labels",
			catAxisTitleColor   : "0088CC",
			catAxisTitleFontSize: 14,
			showCatAxisTitle    : true,

			// Data Labels
			showLabel             : true, // Must be set to true or labels will not be shown
			dataLabelFormatScatter: 'custom', // Can be set to `custom` (default), `customXY`, or `XY`.
		};
		slide.addChart( pptx.charts.SCATTER, arrDataScatterLabels, optsChartScat3 );

		// BOTTOM-RIGHT
		var optsChartScat4 = { x:7.0, y:4.0, w:'45%', h:3 };
		slide.addChart( pptx.charts.SCATTER, arrDataScatter2, optsChartScat4 );
	}

	// SLIDE 15: Bubble Charts -------------------------------------------------------------
	function slide15() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Bubble Charts', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		var arrDataBubble1 = [
			{ name:'X-Axis',    values:[0.3,0.6,0.9,1.2,1.5,1.7] },
			{ name:'Y-Value 1', values:[1.3, 9, 7.5, 2.5, 7.5,  5], sizes:[1,4,2,3,7,4] },
			{ name:'Y-Value 2', values:[  5, 3,   2,   7,   2, 10], sizes:[9,7,9,2,4,8] }
		];
		var arrDataBubble2 = [
			{ name:'X-Axis',   values:[1, 2, 3, 4, 5, 6] },
			{ name:'Airplane', values:[33, 20, 51, 65, 71, 75], sizes:[10,10,12,12,15,20] },
			{ name:'Train',    values:[99, 88, 77, 89, 99, 99], sizes:[20,20,22,22,25,30] },
			{ name:'Bus',      values:[21, 22, 25, 49, 59, 69], sizes:[11,11,13,13,16,21] }
		];

		// TOP-LEFT
		var optsChartBubble1 = { x:0.5, y:0.6, w:'45%', h:3,
			chartColors: ['4477CC','ED7D31'],
			chartColorsOpacity: 40,
			dataBorder: {pt:1, color:'FFFFFF'}
		};
		slide.addText( '.', {x:0.5, y:0.6, w:6.0, h:3.0, fill:'F1F1F1', color:'F1F1F1'} );
		slide.addChart( pptx.charts.BUBBLE, arrDataBubble1, optsChartBubble1 );

		// TOP-RIGHT
		var optsChartBubble2 = { x:7.0, y:0.6, w:'45%', h:3,
			fill: 'f1f1f1',
			showLegend: true,
			legendPos : 'b',

			lineSize  : 8,
			lineSmooth: true,
			lineDataSymbolSize: 12,
			lineDataSymbolLineColor: 'FFFFFF',

			chartColors: [ COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK ],
			chartColorsOpacity: 25
		};
		slide.addChart( pptx.charts.BUBBLE, arrDataBubble2, optsChartBubble2 );

		// BOTTOM-LEFT
		var optsChartBubble3 = { x:0.5, y:4.0, w:'45%', h:3,
			fill: 'f2f9fc',
			catAxisOrientation: 'maxMin',
			valAxisOrientation: 'maxMin',
			showCatAxisTitle: false,
			showValAxisTitle: false,
			valAxisMinVal: 0,
			dataBorder: {pt:2, color:'FFFFFF'},
			dataLabelColor: 'FFFFFF',
			showValue: true
		};
		slide.addChart( pptx.charts.BUBBLE, arrDataBubble1, optsChartBubble3 );

		// BOTTOM-RIGHT
		var optsChartBubble4 = { x:7.0, y:4.0, w:'45%', h:3, lineSize:0 };
		slide.addChart( pptx.charts.BUBBLE, arrDataBubble2, optsChartBubble4 );
	}

	// SLIDE 15: Radar Chart ---------------------------------------------------------------
	function slide16() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Radar Chart', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		var arrDataRegions = [
			{
				name  : 'Region 1',
				labels: ['May', 'June', 'July', 'August', 'September'],
				values: [26, 53, 100, 75, 41]
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

		// TOP-LEFT: Standard
		var optsChartRadar1 = { x:0.5, y:0.6, w:6.0, h:3.0,
			radarStyle: 'standard',
			lineDataSymbol: 'none',
			fill: 'F1F1F1'
		};
		slide.addChart( pptx.charts.RADAR, arrDataRegions, optsChartRadar1 );

		// TOP-RIGHT: Marker
		var optsChartRadar2 = { x:7.0, y:0.6, w:6.0, h:3.0,
			radarStyle: 'marker',
			catAxisLabelColor   : '0000CC',
			catAxisLabelFontFace: 'Courier',
			catAxisLabelFontSize: 12
		};
		slide.addChart( pptx.charts.RADAR, arrDataRegions, optsChartRadar2 );

		// BTM-LEFT: Filled - TITLE and LEGEND
		slide.addText( '.', { x:0.5, y:3.8, w:6.0, h:3.5, fill:'F1F1F1', color:'F1F1F1'} );
		var optsChartRadar3 = { x:0.5, y:3.8, w:6.0, h:3.5,
			radarStyle: 'filled',
			catAxisLabelColor   : 'CC0000',
			catAxisLabelFontFace: 'Helvetica Neue',
			catAxisLabelFontSize: 14,

			showTitle : true,
			titleColor   : '33CF22',
			titleFontFace: 'Helvetica Neue',
			titleFontSize: 16,
			title: 'Sales by Region',

			showLegend : true
		};
		slide.addChart( pptx.charts.RADAR, arrDataHighVals, optsChartRadar3 );

		// BTM-RIGHT: TITLE and LEGEND
		slide.addText( '.', { x:7.0, y:3.8, w:6.0, h:3.5, fill:'F1F1F1', color:'F1F1F1'} );
		var optsChartRadar4 = { x:7.0, y:3.8, w:6.0, h:3.5,
			radarStyle: 'filled',
			chartColors: ['0088CC', '99FFCC'],

			catAxisLabelColor   : '0000CC',
			catAxisLabelFontFace: 'Times',
			catAxisLabelFontSize: 11,
			catAxisLineShow: false,

			showLegend : true,
			legendPos  :  't',
			legendColor: 'FF0000',
			showTitle  : true,
			titleColor : 'FF0000',
			title	  : 'Red Title and Legend'
		};
		slide.addChart( pptx.charts.RADAR, arrDataHighVals, optsChartRadar4 );
	}

	// SLIDE 16: Multi-Type Charts ---------------------------------------------------------
	function slide17() {
		// powerpoint 2016 add secondary category axis labels
		// https://peltiertech.com/chart-with-a-dual-category-axis/

		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Examples: Multi-Type Charts', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		function doStackedLine() {
			// TOP-RIGHT:
			var opts = {
				x: 7.0, y: 0.6, w: 6.0, h: 3.0,
				barDir: 'col',
				barGrouping: 'stacked',
				catAxisLabelColor: '0000CC',
				catAxisLabelFontFace: 'Arial',
				catAxisLabelFontSize: 12,
				catAxisOrientation: 'minMax',
				showLegend: false,
				showTitle: false,
				valAxisMaxVal: 100,
				valAxisMajorUnit: 10
			};

			var labels = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
			var chartTypes = [
				{
					type: pptx.charts.BAR,
					data: [
						{
							name: 'Bottom',
							labels: labels,
							values: [17, 26, 53, 10, 4]
						},
						{
							name: 'Middle',
							labels: labels,
							values: [55, 40, 20, 30, 15]
						},
						{
							name: 'Top',
							labels: labels,
							values: [10, 22, 25, 35, 70]
						}
					],
					options: {
						barGrouping: 'stacked'
					}
				},
				{
					type: pptx.charts.LINE,
					data: [{
						name: 'Current',
						labels: labels,
						values: [25, 35, 55, 10, 5]
					}],
					options: {
						barGrouping: 'standard'
					}
				}
			];
			slide.addChart(chartTypes, opts);
		}

		function doColumnAreaLine() {
			var opts = {
				x: 0.6, y: 0.6, w: 6.0, h: 3.0,
				barDir: 'col',
				catAxisLabelColor: '666666',
				catAxisLabelFontFace: 'Arial',
				catAxisLabelFontSize: 12,
				catAxisOrientation: 'minMax',
				showLegend: false,
				showTitle: false,
				valAxisMaxVal: 100,
				valAxisMajorUnit: 10,

				valAxes:[
					{
						showValAxisTitle: true,
						valAxisTitle: 'Primary Value Axis'
					}, {
						showValAxisTitle: true,
						valAxisTitle: 'Secondary Value Axis',
						valGridLine: 'none'
					}
				],

				catAxes: [
					{
						catAxisTitle: 'Primary Category Axis'
					}, {
						catAxisHidden: true
					}
				]
			};

			var labels = ['April', 'May', 'June', 'July', 'August'];
			var chartTypes = [
				{
					type: pptx.charts.AREA,
					data: [{
						name: 'Current',
						labels: labels,
						values: [1, 4, 7, 2, 3]
					}],
					options: {
						chartColors: ['00FFFF'],
						barGrouping: 'standard',
						secondaryValAxis: !!opts.valAxes,
						secondaryCatAxis: !!opts.catAxes
					}
				}, {
					type: pptx.charts.BAR,
					data: [{
						name: 'Bottom',
						labels: labels,
						values: [17, 26, 53, 10, 4]
					}],
					options: {
						chartColors: ['0000FF'],
						barGrouping: 'stacked'
					}
				}, {
					type: pptx.charts.LINE,
					data: [{
						name: 'Current',
						labels: labels,
						values: [5, 3, 2, 4, 7]
					}],
					options: {
						barGrouping: 'standard',
						secondaryValAxis: !!opts.valAxes,
						secondaryCatAxis: !!opts.catAxes
					}
				}
			];
			slide.addChart(chartTypes, opts);
		}

		function doStackedDot() {
			// BOT-LEFT:
			var opts = {
				x: 0.6, y: 4.0, w: 6.0, h: 3.0,
				barDir: 'col',
				barGrouping: 'stacked',
				catAxisLabelColor: '999999',
				catAxisLabelFontFace: 'Arial',
				catAxisLabelFontSize: 14,
				catAxisOrientation: 'minMax',
				showLegend: false,
				showTitle: false,
				valAxisMaxVal: 100,
				valAxisMinVal: 0,
				valAxisMajorUnit: 20,

				lineSize: 0,
				lineDataSymbolSize: 20,
				lineDataSymbolLineSize: 2,
				lineDataSymbolLineColor: 'FF0000',

				//dataNoEffects: true,

				valAxes:[
					{
						showValAxisTitle: true,
						valAxisTitle: 'Primary Value Axis'
					}, {
						showValAxisTitle: true,
						valAxisTitle: 'Secondary Value Axis',
						catAxisOrientation  : 'maxMin',
						valAxisMajorUnit: 1,
						valAxisMaxVal: 10,
						valAxisMinVal: 1,
						valGridLine: 'none'
					}
				],
				catAxes: [
					{
						catAxisTitle: 'Primary Category Axis'
					}, {
						catAxisHidden: true
					}

				]
			};

			var labels = ['Q1', 'Q2', 'Q3', 'Q4', 'OT'];
			var chartTypes = [
				{
					type: pptx.charts.BAR,
					data: [{
						name: 'Bottom',
						labels: labels,
						values: [17, 26, 53, 10, 4]
					},
						{
							name: 'Middle',
							labels: labels,
							values: [55, 40, 20, 30, 15]
						},
						{
							name: 'Top',
							labels: labels,
							values: [10, 22, 25, 35, 70]
						}],
					options: {
						barGrouping: 'stacked'
					}
				}, {
					type: pptx.charts.LINE,
					data: [{
						name: 'Current',
						labels: labels,
						values: [5, 3, 2, 4, 7]
					}],
					options: {
						barGrouping: 'standard',
						secondaryValAxis: !!opts.valAxes,
						secondaryCatAxis: !!opts.catAxes,
						chartColors: ['FFFF00']
					}
				}
			];
			slide.addChart(chartTypes, opts);
		}

		function doBarCol() {
			// BOT-RGT:
			var opts = {
				x: 7, y: 4.0, w: 6.0, h: 3.0,
				barDir: 'col',
				barGrouping: 'stacked',
				catAxisLabelColor: '999999',
				catAxisLabelFontFace: 'Arial',
				catAxisLabelFontSize: 14,
				catAxisOrientation: 'minMax',
				showLegend: false,
				showTitle: false,
				valAxisMaxVal: 100,
				valAxisMinVal: 0,
				valAxisMajorUnit: 20,
				valAxes:[
					{
						showValAxisTitle: true,
						valAxisTitle: 'Primary Value Axis'
					}, {
						showValAxisTitle: true,
						valAxisTitle: 'Secondary Value Axis',
						catAxisOrientation  : 'maxMin',
						valAxisMajorUnit: 1,
						valAxisMaxVal: 10,
						valAxisMinVal: 1,
						valGridLine: 'none'
					}
				],
				catAxes: [
					{
						catAxisTitle: 'Primary Category Axis'
					}, {
						catAxisHidden: true
					}

				]
			};

			var labels = ['Q1', 'Q2', 'Q3', 'Q4', 'OT'];
			var chartTypes = [
				{
					type: pptx.charts.BAR,
					data: [{
						name: 'Bottom',
						labels: labels,
						values: [17, 26, 53, 10, 4]
					},
						{
							name: 'Middle',
							labels: labels,
							values: [55, 40, 20, 30, 15]
						},
						{
							name: 'Top',
							labels: labels,
							values: [10, 22, 25, 35, 70]
						}],
					options: {
						barGrouping: 'stacked'
					}
				}, {
					type: pptx.charts.BAR,
					data: [{
						name: 'Current',
						labels: labels,
						values: [5, 3, 2, 4, 7]
					}],
					options: {
						barDir: 'bar',
						barGrouping: 'standard',
						secondaryValAxis: !!opts.valAxes,
						secondaryCatAxis: !!opts.catAxes
					}
				}
			];
			slide.addChart(chartTypes, opts);
		}

		function readmeExample() {
			// for testing - not rendered in demo
			var labels = ['Q1', 'Q2', 'Q3', 'Q4', 'OT'];
			var chartTypes = [
				{
					type: pptx.charts.BAR,
					data: [{
						name: 'Projected',
						labels: labels,
						values: [17, 26, 53, 10, 4]
					}],
					options: {
						barDir: 'col'
					}
				}, {
					type: pptx.charts.LINE,
					data: [{
						name: 'Current',
						labels: labels,
						values: [5, 3, 2, 4, 7]
					}],
					options: {
						secondaryValAxis: true,
						secondaryCatAxis: true
					}
				}
			];
			var multiOpts = {
				x:1.0, y:1.0, w:6, h:6,
				valAxisMaxVal: 100,
				valAxisMinVal: 0,
				valAxisMajorUnit: 20,
				valAxes:[
					{
						showValAxisTitle: true,
						valAxisTitle: 'Primary Value Axis'
					}, {
						showValAxisTitle: true,
						valAxisTitle: 'Secondary Value Axis',
						valAxisMajorUnit: 1,
						valAxisMaxVal: 10,
						valAxisMinVal: 1,
						valGridLine: 'none'
					}
				],
				catAxes: [
					{
						catAxisTitle: 'Primary Category Axis'
					}, {
						catAxisHidden: true
					}

				]
			};

			slide.addChart(chartTypes, multiOpts);
		}

		doBarCol();
		doStackedDot();
		doColumnAreaLine();
		doStackedLine();
		//readmeExample();
	}

	// SLIDE 17: Charts Options: Shadow, Transparent Colors --------------------------------
	function slide18() {
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-charts.html');
		slide.addTable( [ [{ text:'Chart Options: Shadow, Transparent Colors', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		var arrDataRegions = [{
			name  : 'Region 2',
			labels: ['April', 'May', 'June', 'July', 'August'],
			values: [0, 30, 53, 10, 25]
		}, {
			name  : 'Region 3',
			labels: ['April', 'May', 'June', 'July', 'August'],
			values: [17, 26, 53, 100, 75]
		}, {
			name  : 'Region 4',
			labels: ['April', 'May', 'June', 'July', 'August'],
			values: [55, 43, 70, 90, 80]
		}, {
			name  : 'Region 5',
			labels: ['April', 'May', 'June', 'July', 'August'],
			values: [55, 43, 70, 90, 80]
		}];
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
		var single = [{
			name  : 'Texas',
			labels: ['Apartment', 'Townhome', 'Duplex', 'House', 'Big House'],
			values: [1400, 2000, 2500, 3000, 3800]
		}];

		// TOP-LEFT: H/bar
		var optsChartBar1 = { x:0.5, y:0.6, w:6.0, h:3.0,
			showTitle: true,
			title: 'Large blue shadow',
			barDir: 'bar',
			barGrouping: 'standard',
			dataLabelColor   : 'FFFFFF',
			showValue        : true,
			shadow: {
				type: 'outer',
				blur: 10,
				offset: 5,
				angle: 45,
				color: '0059B1',
				opacity: 1
			}
		};

		var pieOptions = { x:7.0, y:0.6, w:6.0, h:3.0,
			showTitle: true,
			title: 'Rotated cyan shadow',
			dataLabelColor   : 'FFFFFF',
			shadow: {
				type: 'outer',
				blur: 10,
				offset: 5,
				angle: 180,
				color: '00FFFF',
				opacity: 1
			}
		};

		// BTM-LEFT: H/bar - 100% layout without axis labels
		var optsChartBar3 = { x:0.5, y:3.8, w:6.0, h:3.5,
			showTitle: true,
			title: 'No shadow, transparent colors',
			barDir     : 'bar',
			barGrouping: 'stacked',
			chartColors: ['transparent', '5DA5DA', 'transparent', 'FAA43A'],
			shadow: 'none'
		};

		// BTM-RIGHT: V/col - TITLE and LEGEND
		var optsChartBar4 = { x:7.0, y:3.8, w:6.0, h:3.5,
			barDir: 'col',
			barGrouping: 'stacked',
			showTitle: true,
			title: 'Red glowing shadow',
			catAxisLabelColor   : '0000CC',
			catAxisLabelFontFace: 'Times',
			catAxisLabelFontSize: 12,
			catAxisOrientation  : 'minMax',
			chartColors: ['5DA5DA','FAA43A'],
			shadow: {
				type: 'outer',
				blur: 20,
				offset: 1,
				angle: 90,
				color: 'A70000',
				opacity: 1
			}
		};

		slide.addChart( pptx.charts.BAR, single, optsChartBar1 );
		slide.addChart( pptx.charts.PIE, dataChartPieStat, pieOptions );
		slide.addChart( pptx.charts.BAR, arrDataRegions, optsChartBar3 );
		slide.addChart( pptx.charts.BAR, arrDataHighVals, optsChartBar4 );
	}

	// RUN ALL SLIDE DEMOS -----
	slide1();
	slide2();
	slide3();
	slide4();
	slide5();
	slide6();
	slide7();
	slide8();
	slide9();
	slide10();
	slide11();
	slide12();
	slide13();
	slide14();
	slide15();
	slide16();
	slide17();
}

function genSlides_Image(pptx) {
	// NOTE:
	// Images can be pre-encoded into base64, so they do not have to be on the webserver etc. (saves generation time and resources!)
	// This also has the benefit of being able to be any type (path:images can only be exported as PNG)
	// Image source: either `data` or `path` is required

	// SLIDE 1: Image Types -----------------------------------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html');
		slide.slideNumber({ x:'50%', y:'95%', color:'0088CC' });
		slide.addTable( [ [{ text:'Image Examples: Misc Image Types', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// TOP: 1
		slide.addText('Type: Animated GIF', { x:0.5, y:0.6, w:2.5, h:0.4, color:'0088CC' });
		slide.addImage({ x:1.0, y:1.1, w:1.5, h:1.5, path:(NODEJS ? gPaths.gifAnimTrippy.path.replace(/http.+\/examples/, '../examples') : gPaths.gifAnimTrippy.path) });
		slide.addText('(use slide Show)', { x:1.0, y:2.7, w:1.5, h:0.3, color:'696969', fill:'FFFCCC', align:'c', fontSize:10 });

		// TOP: 2
		slide.addText('Type: GIF', { x:4.35, y:0.6, w:1.4, h:0.4, color:'0088CC' });
		slide.addImage({ x:4.4, y:1.05, w:1.2, h:1.2, path:(NODEJS ? gPaths.ccDjGif.path.replace(/http.+\/examples/, '../examples') : gPaths.ccDjGif.path) });

		// TOP: 3
		slide.addText('Type: base64 PNG', { x:7.2, y:0.6, w:2.4, h:0.4, color:'0088CC' });
		slide.addImage({ x:7.87, y:1.1, w:1.0, h:1.0, data:checkGreen });

		// TOP: 4
		slide.addText('Image Hyperlink', { x:10.9, y:0.6, w:2.2, h:0.4, color:'0088CC' });
		slide.addImage({
			x:11.54, y:1.2, w:0.8, h:0.8,
			data: svgHyperlinkImage,
			hyperlink: { url:'https://github.com/gitbrent/pptxgenjs', tooltip:'Visit Homepage' }
		});

		// BOTTOM-LEFT:
		slide.addText('Type: JPG', { x:0.5, y:3.3, w:4.5, h:0.4, color:'0088CC' });
		slide.addImage({ path:gPaths.ccCopyRemix.path, x:0.5, y:3.8, w:3.0, h:3.07 });

		// BOTTOM-CENTER:
		slide.addText('Type: PNG', { x:5.1, y:3.3, w:4.0, h:0.4, color:'0088CC' });
		slide.addImage({ path:gPaths.wikimedia1.path, x:5.1, y:3.8, w:3.0, h:2.78 });

		// BOTTOM-RIGHT:
		slide.addText('Type: SVG', { x:9.5, y:3.3, w:4.0, h:0.4, color:'0088CC' });
		slide.addImage({ path:gPaths.wikimedia_svg1.path, x:9.5, y:3.8, w:2.0, h:2.0 }); // TEST: `path`
		slide.addImage({ data:svgBase64, x:11.1, y:5.1, w:1.5, h:1.5 }); // TEST: `data`
		slide.addText('(not supported in Node)', { x:9.1, y:6.8, w:3.5, h:0.3, color:'696969', fill:'FFFCCC', align:'c', fontSize:10 });

		// TEST: Ensure framework corrects for missing all header
		// (Please **DO NOT** pass base64 data without the header! This is a JUNK TEST!)
		//slide.addImage({ x:5.2, y:2.6, w:0.8, h:0.8, data:'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAAjcAAAI3AGf6F88AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAANVQTFRF////JLaSIJ+AIKqKKa2FKLCIJq+IJa6HJa6JJa6IJa6IJa2IJa6IJa6IJa6IJa6IJa6IJa6IJq6IKK+JKK+KKrCLLrGNL7KOMrOPNrSRN7WSPLeVQrmYRLmZSrycTr2eUb6gUb+gWsKlY8Wqbsmwb8mwdcy0d8y1e863g9G7hdK8htK9i9TAjNTAjtXBktfEntvKoNzLquDRruHTtePWt+TYv+fcx+rhyOvh0e7m1e/o2fHq4PTu5PXx5vbx7Pj18fr49fv59/z7+Pz7+f38/P79/f7+dNHCUgAAABF0Uk5TAAcIGBktSYSXmMHI2uPy8/XVqDFbAAABB0lEQVQ4y42T13qDMAyFZUKMbebp3mmbrnTvlY60TXn/R+oFGAyYzz1Xx/wylmWJqBLjUkVpGinJGXXliwSVEuG3sBdkaCgLPJMPQnQUDmo+jGFRPKz2WzkQl//wQvQoLPII0KuAiMjP+gMyn4iEFU1eAQCCiCU2fpCfFBVjxG18f35VOk7Swndmt9pKUl2++fG4qL2iqMPXpi8r1SKitDDne/rT8vPbRh2d6oC7n6PCLNx/bsEM0Edc5DdLAHD9tWueF9VJjmdP68DZ77iRkDKuuT19Hx3mx82MpVmo1Yfv+WXrSrxZ6slpiyes77FKif88t7Nh3C3nbFp327sHxz167uHtH/8/eds7gGsUQbkAAAAASUVORK5CYII=' });
		// NEGATIVE-TEST:
		//slide.addImage({ data:'https://cdn.rawgit.com/gitbrent/PptxGenJS/v2.1.0/examples/images/doh_this_isnt_base64_data.gif',  x:0.5, y:0.5, w:1.0, h:1.0 });
	}

	// SLIDE 2: Image URLs -----------------------------------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html');
		slide.slideNumber({ x:'50%', y:'95%', color:'0088CC' });
		slide.addTable( [ [{ text:'Image Examples: Image URLs', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// TOP-LEFT:
		var objCodeEx1 = { x:0.5, y:0.6, w:6.0, h:0.6 };
		Object.keys(gOptsCode).forEach(function(key){ objCodeEx1[key] = gOptsCode[key] });
		slide.addText('path:"'+gPaths.ccLogo.path+'"', objCodeEx1);
		slide.addImage({ path:gPaths.ccLogo.path, x:0.5, y:1.35, h:2.5, w:3.33 });

		// TOP-RIGHT:
		var objCodeEx2 = { x:6.9, y:0.6, w:6.0, h:0.6 };
		Object.keys(gOptsCode).forEach(function(key){ objCodeEx2[key] = gOptsCode[key] });
		slide.addText('path:"'+gPaths.wikimedia2.path+'"', objCodeEx2);
		slide.addImage({ path:gPaths.wikimedia2.path, x:6.9, y:1.35, h:2.5, w:3.27 });

		// BTM-LEFT:
		var objCodeEx3 = { x:0.5, y:4.2, w:12.4, h:0.8 };
		Object.keys(gOptsCode).forEach(function(key){ objCodeEx3[key] = gOptsCode[key] });
		slide.addText('// Test: URL variables, plus more than one ".jpg"\npath:"'+gPaths.chicagoBean.path+'"', objCodeEx3);
		slide.addImage({ path:gPaths.chicagoBean.path, x:0.5, y:5.1, w:2.56, h:1.92 });

		// BOTTOM-CENTER:
		if ( typeof window !== 'undefined' && window.location.href.indexOf('gitbrent') > 0 ) {
			// TEST USING RELATIVE PATHS/LOCAL FILES (OFFICE.COM)
			slide.addText('Type: PNG (path:"../images")', { x:6.6, y:2.7, w:4.5, h:0.4, color:'CC0033' });
			slide.addImage({ path:(NODEJS ? gPaths.ccLicenseComp.path.replace(/http.+\/examples/, '../examples') : gPaths.ccLicenseComp.path), x:6.6, y:3.2, w:6.3, h:3.7 });
		}
	}

	// SLIDE 3: Image Sizing -----------------------------------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-images.html');
		slide.slideNumber({ x:'50%', y:'95%', w:1, h:1, color:'0088CC' });
		slide.addTable( [ [{ text:'Image Examples: Image Sizing/Rounding', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// TOP: 1
		slide.addText('Sizing: None',      { x:0.5, y:0.6, w:3.0, h:0.3, color:'0088CC' });
		slide.addImage({ data:LOGO_STARLABS, x:0.5, y:1.1, w:5.0, h:2.5 });

		// TOP: 2
		slide.addText("Sizing: `contain`", { x:0.6, y:4.25, w:3.0, h:0.3, color:'0088CC' });
		slide.addImage({ data:LOGO_STARLABS, x:0.6, y:4.65, w:5.0, h:2.5, sizing:{ type:'contain', w:3, h:1.5 } });

		// TOP: 3
		slide.addText('Sizing: `cover`',   { x:5.3, y:4.25, w:3.0, h:0.3, color:'0088CC' });
		slide.addImage({ data:LOGO_STARLABS, x:5.3, y:4.65, w:3.0, h:1.5, sizing:{ type:'cover', w:4, h:1.88 } });

		// TOP: 4
		slide.addText('Sizing: `crop`',    { x:10.0, y:4.25, w:3.0, h:0.3, color:'0088CC' });
		slide.addImage({ data:LOGO_STARLABS, x:10.0, y:4.65, w:5.0, h:2.5, sizing:{ type:'crop', w:3, h:1.5, x:0.5, y:0.5 } });

		// TOP-RIGHT:
		slide.addText('Rounding: `true`',  { x:10.0, y:0.60, w:3.0, h:0.3, color:'0088CC' });
		slide.addImage({
			path:(NODEJS ? gPaths.ccLogo.path.replace(/http.+\/examples/, '../examples') : gPaths.ccLogo.path),
			x:9.9, y:1.1, w:2.5, h:2.5,
			rounding:true
		});
	}
}

function genSlides_Media(pptx) {
	// SLIDE 1: Video and YouTube -----------------------------------------------------------------------------------
	var slide1 = pptx.addNewSlide();
	slide1.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html');
	slide1.addTable( [ [{ text:'Media: Misc Video Formats; YouTube', opts:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

	slide1.addText('Video: m4v', { x:0.5, y:0.6, w:4.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:0.5, y:1.0, w:4.00, h:2.27, type:'video', path:(NODEJS ? gPaths.sample_m4v.path.replace(/http.+\/examples/, '../examples') : gPaths.sample_m4v.path) });

	slide1.addText('Video: mpg', { x:5.5, y:0.6, w:3.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:5.5, y:1.0, w:3.00, h:2.05, type:'video', path:(NODEJS ? gPaths.sample_mpg.path.replace(/http.+\/examples/, '../examples') : gPaths.sample_mpg.path) });

	slide1.addText('Video: mov', { x:9.4, y:0.6, w:3.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:9.4, y:1.0, w:3.00, h:1.71, type:'video', path:(NODEJS ? gPaths.sample_mov.path.replace(/http.+\/examples/, '../examples') : gPaths.sample_mov.path) });

	slide1.addText('Video: mp4', { x:0.5, y:3.6, w:4.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:0.5, y:4.0, w:4.00, h:3.00, type:'video', path:(NODEJS ? gPaths.sample_mp4.path.replace(/http.+\/examples/, '../examples') : gPaths.sample_mp4.path) });

	slide1.addText('Video: avi', { x:5.5, y:3.6, w:3.00, h:0.4, color:'0088CC' });
	slide1.addMedia({ x:5.5, y:4.0, w:3.00, h:2.25, type:'video', path:(NODEJS ? gPaths.sample_avi.path.replace(/http.+\/examples/, '../examples') : gPaths.sample_avi.path) });

	// NOTE: Only generated on Node as I dont want everyone who downloads and runs this to be greated with an error!
	if ( !NODEJS && $ && $('#chkYoutube').prop('checked') ) {
		slide1.addText('Online: YouTube', { x:9.4, y:3.6, w:3.00, h:0.4, color:'0088CC' });
		// Provide the usual options (locations and size), then pass the embed code from YouTube (it's on every video page)
		slide1.addMedia({ x:9.4, y:4.0, w:3.00, h:2.25, type:'online', link:'https://www.youtube.com/embed/Dph6ynRVyUc' });

		slide1.addText(
			'**NOTE** YouTube videos will issue a content warning in older desktop PPT (they only work in PPT Online/Desktop v16+)',
			{ shape:pptx.shapes.RECTANGLE, x:0.0, y:7.0, w:'100%', h:0.53, fill:'FFF000', align:'c', fontSize:12 }
		);
	}

	// SLIDE 2: Audio -----------------------------------------------------------------------------------
	var slide2 = pptx.addNewSlide();
	slide2.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-media.html');
	slide2.addTable( [ [{ text:'Media: Misc Audio Formats', opts:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

	slide2.addText('Audio: mp3', { x:0.5, y:0.6, w:4.00, h:0.4, color:'0088CC' });
	slide2.addMedia({ x:0.5, y:1.0, w:4.00, h:0.3, type:'audio', path:(NODEJS ? gPaths.sample_mp3.path.replace(/http.+\/examples/, '../examples') : gPaths.sample_mp3.path) });

	slide2.addText('Audio: wav', { x:0.5, y:2.6, w:4.00, h:0.4, color:'0088CC' });
	slide2.addMedia({ x:0.5, y:3.0, w:4.00, h:0.3, type:'audio', path:(NODEJS ? gPaths.sample_wav.path.replace(/http.+\/examples/, '../examples') : gPaths.sample_wav.path) });

	if ( typeof window !== 'undefined' && window.location.href.indexOf('gitbrent') > 0 ) {
		// TEST USING LOCAL FILES (OFFICE.COM)
		slide2.addText('Audio: MP3 (path:"../media")', { x:0.5, y:4.6, w:4.0, h:0.4, color:'0088CC' });
		slide2.addMedia({ x:0.5, y:5.0, w:4.0, h:0.3, type:'audio', path:'../SiteAssets/pptxgenjs/examples/media/sample.mp3' });
	}
}

function genSlides_Shape(pptx) {
	// SLIDE 1: Misc Shape Types (no text)
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-shapes.html');
	slide.addTable( [ [{ text:'Shape Examples 1: Misc Shape Types (no text)', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

	//slide.addShape(pptx.shapes.RECTANGLE,         { x:0.5, y:0.8, w:12.5,h:0.5, fill:'F9F9F9' });
	slide.addShape(pptx.shapes.RECTANGLE,         { x:0.5, y:0.8, w:1.5, h:3.0, fill:'FF0000' });
	slide.addShape(pptx.shapes.RECTANGLE,         { x:3.0, y:0.7, w:1.5, h:3.0, fill:'F38E00', rotate:45 });
	slide.addShape(pptx.shapes.OVAL,              { x:5.4, y:0.8, w:3.0, h:1.5, fill:{ type:'solid', color:'0088CC', alpha:25 } });
	slide.addShape(pptx.shapes.OVAL,              { x:7.7, y:1.4, w:3.0, h:1.5, fill:{ type:'solid', color:'FF00CC', alpha:50 }, rotate:90 });
	slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:10 , y:2.5, w:3.0, h:1.5, r:0.2, fill:'00FF00', line:'000000', lineSize:1 });
	//
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:4.4, w:5.0, h:0.0, line:'FF0000', lineSize:1 });
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:4.8, w:5.0, h:0.0, line:'FF0000', lineSize:2, lineHead:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:5.2, w:5.0, h:0.0, line:'FF0000', lineSize:3, lineTail:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.2, y:5.6, w:5.0, h:0.0, line:'FF0000', lineSize:4, lineHead:'triangle', lineTail:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:5.7, y:3.3, w:2.5, h:0, rotate:(360-45) }); // DIAGONAL Line // TEST: (missing `line`, `lineSize`)
	//
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:0.4, y:4.3, w:6.0, h:3.0, fill:'0088CC', line:'000000', lineSize:3 });
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:7.0, y:4.3, w:6.0, h:3.0, fill:'0088CC', line:'000000', flipH:true });

	// SLIDE 2: Misc Shape Types with Text
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-shapes.html');
	slide.addTable( [ [{ text:'Shape Examples 2: Misc Shape Types (with text)', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

	slide.addText('RECTANGLE',                  { shape:pptx.shapes.RECTANGLE, x:0.5, y:0.8, w:1.5, h:3.0, fill:'FF0000', align:'c', fontSize:14 });
	slide.addText('RECTANGLE (rotate:45)',      { shape:pptx.shapes.RECTANGLE, x:3.0, y:0.7, w:1.5, h:3.0, fill:'F38E00', rotate:45, align:'c', fontSize:14 });
	slide.addText('OVAL (alpha:25)',            { shape:pptx.shapes.OVAL,      x:5.4, y:0.8, w:3.0, h:1.5, fill:{ type:'solid', color:'0088CC', alpha:25 }, align:'c', fontSize:14 });
	slide.addText('OVAL (rotate:90, alpha:50)', { shape:pptx.shapes.OVAL,      x:7.7, y:1.4, w:3.0, h:1.5, fill:{ type:'solid', color:'FF00CC', alpha:50 }, rotate:90, align:'c', fontSize:14 });
	slide.addText('ROUNDED-RECTANGLE\nlineDash:dash\nrectRadius:10', { shape:pptx.shapes.ROUNDED_RECTANGLE, x:10 , y:2.5, w:3.0, h:1.5, r:0.2, fill:'00FF00', align:'c', fontSize:14, line:'000000', lineSize:1, lineDash:'dash', rectRadius:10 });
	//
	slide.addText('LINE size=1',     { shape:pptx.shapes.LINE, align:'c', x:4.15, y:4.40, w:5, h:0, line:'FF0000', lineSize:1, lineDash:'lgDash' });
	slide.addText('LINE size=2',     { shape:pptx.shapes.LINE, align:'l', x:4.15, y:4.80, w:5, h:0, line:'FF0000', lineSize:2, lineTail:'triangle' });
	slide.addText('LINE size=3',     { shape:pptx.shapes.LINE, align:'r', x:4.15, y:5.20, w:5, h:0, line:'FF0000', lineSize:3, lineHead:'triangle' });
	slide.addText('LINE size=4',     { shape:pptx.shapes.LINE, align:'c', x:4.15, y:5.60, w:5, h:0, line:'FF0000', lineSize:4, lineHead:'triangle', lineTail:'triangle' });
	slide.addText('DIAGONAL',        { shape:pptx.shapes.LINE, valign:'b', x:5.7, y:3.3, w:2.5, h:0, lineSize:2, rotate:(360-45) }); // TEST: (missing `line`)
	//
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:0.4, y:4.3, w:6, h:3, fill:'0088CC', line:'000000', lineSize:3 });
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE, align:'c', x:7.0, y:4.3, w:6, h:3, fill:'0088CC', line:'000000', flipH:true });
}

function genSlides_Text(pptx) {
	// SLIDE 1: Multi-Line Formatting, Hyperlinks, Text Shadow, Line-Break
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html');
		// NOTE: not using `gOptsTabOpts` here to test legacy `cx`
		slide.addTable( [ [{ text:'Text Examples: Multi-Line Formatting, Hyperlinks, Text Shadow, Line-Break', options:gOptsTextL },gOptsTextR] ], { x:0.5, y:0.13, cx:12.6 } );

		// LEFT COLUMN ------------------------------------------------------------

		// 1: Multi-Line Formatting
		slide.addText("Word-Level Formatting:", { x:0.5, y:0.5, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text:'1st\nline',options:{ fontSize:24, fontFace:'Courier New', color:'99ABCC', align:'r', breakLine:true } },
				{ text:'2nd line', options:{ fontSize:36, fontFace:'Arial',       color:'FFFF00', align:'c', breakLine:true } },
				{ text:'3rd line', options:{ fontSize:48, fontFace:'Verdana',     color:'0088CC', align:'l' } },
				{ text:'4th line', options:{ fontSize:38, fontFace:'Arial',       color:'FFFF00', align:'c', strike:true } },
				{ text:'5th\nline',options:{ fontSize:36, fontFace:'Courier New', color:'99ABCC', align:'r' } }
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
			{ x:0.5, y:5.35, w:6.0, h:0.6, margin:0.1, fill:'F1F1F1', fontSize:14 }
		);

		// 4: Text Effects: Shadow
		var shadowOpts = { type:'outer', color:'696969', blur:3, offset:10, angle:45, opacity:0.6 };
		slide.addText("Text Shadow:", { x:0.5, y:6.2, w:'40%', h:0.38, color:'0088CC' });
		slide.addText("Text Shadow", { x:0.5, y:6.5, w:'40%', h:0.6, fontSize:32, color:'0088cc', shadow:shadowOpts });

		// RIGHT COLUMN ------------------------------------------------------------

		// 2: Line-Break Test
		slide.addText("Line-Breaks:", { x:7.5, y:0.5, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			'***Line-Break/Multi-Line Test***\n\nFirst line\nSecond line\nThird line',
			{ x:7.5, y:0.85, w:5.25, h:1.6, valign:'middle', align:'ctr', color:'6c6c6c', fontSize:16, fill:'F2F2F2', line:{pt:'2',color:'C7C7C7'} }
		);

		slide.addText("Line-Spacing (text):", { x:7.5, y:2.6, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			'1st Line\n2nd Line (40pt)',
			{ x:7.5, y:2.95, w:5.25, h:1.25, valign:'m', align:'c', fill:'f1f1f1', color:'363636', lineSpacing:40 }
		);

		slide.addText("Line-Spacing (bullets):", { x:7.5, y:4.45, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text:'Line Spacing\n35pt', options:{ fontSize:24, bullet:true, color:'99ABCC', lineSpacing:35 } }
			],
			{ x:7.5, y:4.85, w:5.25, h:1.15, margin:0.1, fill:'f1f1f1' }
		);

		slide.addText("Text Outline:", { x:7.5, y:6.2, w:'40%', h:0.38, color:'0088CC' });
		slide.addText("Text Outline", { x:7.5, y:6.5, w:'40%', h:0.6, fontSize:32, color:'0088cc', outline:{size:1.5, color:'00DD00'} });
	}

	// SLIDE 2: Bullets
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html');
		slide.addTable( [ [{ text:'Text Examples: Bullets', options:gOptsTextL },gOptsTextR] ], { x:0.5, y:0.13, cx:12.5 } );

		// LEFT COLUMN ------------------------------------------------------------

		// 1: Bullets with indent levels
		slide.addText("Bullet Indent-Levels:", { x:0.5, y:0.5, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text:'Bullet one',     options:{ fontSize:24, bullet:true, color:'99ABCC'                } },
				{ text:'Indent-Level 1', options:{ fontSize:32, bullet:true, color:'FFFF00', indentLevel:1 } },
				{ text:'Indent-Level 2', options:{ fontSize:42, bullet:true, color:'0088CC', indentLevel:2 } },
				{ text:'Indent-Level 3', options:{ fontSize:48, bullet:true, color:'CC88BB', indentLevel:3 } },
				{ text:'Indent-Level 3', options:{ fontSize:48, bullet:true, color:'CC88BB', indentLevel:3 } },
				{ text:'Indent-Level 2', options:{ fontSize:42, bullet:true, color:'0088CC', indentLevel:2 } },
				{ text:'Indent-Level 2', options:{ fontSize:42, bullet:true, color:'0088CC', indentLevel:2 } },
				{ text:'Indent-Level 1', options:{ fontSize:32, bullet:true, color:'FFFF00', indentLevel:1 } },
				{ text:'Bullet no indent', options:{ fontSize:24, bullet:true, color:'99ABCC'                } },
				{ text:'Bullet Last',    options:{ fontSize:24, bullet:true, color:'99ABCC'                } }
			],
			{ x:0.5, y:1.0, w:6.25, h:6.0, fill:'373737' }
		);

		// 4: Regular bullets
		slide.addText("Bullets:", { x:7.5, y:0.65, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(12345                  , { x:8.0, y:1.05, w:'30%', h:0.5, color:'0000DE', fontFace:"Courier New", bullet:true });
		slide.addText('String (number above)', { x:8.0, y:1.35, w:'30%', h:0.5, color:'00AA00', bullet:true });

		// 5: Bullets: Text With Line-Breaks
		slide.addText("Bullets with line-breaks:", { x:7.5, y:2.1, w:'40%', h:0.38, color:'0088CC' });
		slide.addText('Line 1\nLine 2\nLine 3', { x:8.0, y:2.6, w:'30%', h:1.0, color:'393939', fontSize:16, fill:'F2F2F2', bullet:{type:'number'} });

		// 6: Bullets: With group of {text}
		slide.addText("Bullet with {text} objects:", { x:7.5, y:4.0, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text: 'big red words... ', options:{fontSize:24, color:'FF0000'} },
				{ text: 'some green words.', options:{fontSize:16, color:'00FF00'} }
			],
			{ x:8.0, y:4.4, w:5.0, h:0.5, margin:0.1, fontFace:'Arial', bullet:{code:'25BA'} }
		);

		// 7: Bullets: Within a {text} object
		slide.addText("Bullet within {text} objects:", { x:7.5, y:5.3, w:'40%', h:0.38, color:'0088CC' });
		slide.addText(
			[
				{ text:'I am a text object with bullets..', options:{bullet:{code:'25BA'}, color:'CC0000'} },
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
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html');
		// Slide colors: bkgd/fore
		slide.back = '030303';
		slide.color = '9F9F9F';
		// Title
		slide.addTable( [ [{ text:'Text Examples: Text alignment, percent x/y, etc.', options:gOptsTextL },gOptsTextR] ], gOptsTabOpts );

		// Actual Textbox shape (can have any Height, can wrap text, etc.)
		slide.addText([
				{ text:'Textbox align (ctr/ctr)', options:{ fontSize:32 } },
				{ text:'Character Spacing 16', options:{ fontSize:16, charSpacing:16 } }
			],
			{ x:0.5, y:0.75, w:8.5, h:2.5, color:'FFFFFF', fill:'0000FF', valign:'c', align:'c', isTextBox:true }
		);
		slide.addText(
			[{ text:'(top/lft)', options:{ fontSize:12 } }, { text:'Textbox', options:{ bold:true } }],
			{ x:10, y:0.75, w:3.0, h:1.0, color:'FFFFFF', fill:'00CC00', valign:'t', align:'l', margin:15 }
		);
		slide.addText(
			[{ text:'Textbox' }, { text:'(btm/rgt)', options:{ fontSize:12 } }],
			{ x:10, y:2.25, w:3.0, h:1.0, color:'FFFFFF', fill:'FF0000', valign:'b', align:'r', margin:0 }
		);

		slide.addText('^ (50%/50%)', { x:'50%', y:'50%', w:2 });

		slide.addText('Plain x/y coords', { x:10, y:4.35, w:3 });

		slide.addText('Escaped chars: \' " & < >', { x:10, y:5.35, w:3 });

		slide.addText(
			[
				{ text:'Sub'},
				{ text:'Subscript', options:{ subscript:true } },
				{ text:' // Super'},
				{ text:'Superscript', options:{ superscript:true } }
			],
			{ x:10, y:6.3, w:3.3 }
		);

		// TEST: using {option}: Add text box with multiline options:
		slide.addText(
			[
				{ text:'word-level\nformatting', options:{ fontSize:32, fontFace:'Courier New', color:'99ABCC', align:'r', breakLine:true } },
				{ text:'...in the same textbox', options:{ fontSize:48, fontFace:'Arial', color:'FFFF00', align:'c' } }
			],
			{ x:0.5, y:4.3, w:8.5, h:2.5, margin:0.1, fill:'232323' }
		);

		var objOptions = {
			x:0, y:7, w:'100%', h:0.5, align:'c',
			fontFace:'Arial', fontSize:24, color:'00EC23', bold:true, italic:true, underline:true, margin:0, isTextBox:true
		};
		slide.addText('Text: Arial, 24, green, bold, italic, underline, margin:0', objOptions);
	}

	// SLIDE 4: Scheme Colors
	{
		var slide = pptx.addNewSlide();
		slide.addNotes('API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-text.html');
		slide.addTable( [ [{ text:'Text Examples: Scheme Colors', options:gOptsTextL },gOptsTextR] ], { x:0.5, y:0.13, cx:12.5 } );

		// MISC ------------------------------------------------------------

		slide.addText("TEXT1 / BACKGROUND2", { x:0.5, y:0.7, w:6.0, h:2.0, color:pptx.colors.TEXT1, fill:pptx.colors.BACKGROUND2 });
		slide.addText("TEXT2 / BACKGROUND1", { x:7.0, y:0.7, w:6.0, h:2.0, color:pptx.colors.TEXT2, fill:pptx.colors.BACKGROUND1 });

		slide.addText("ACCENT1", { x:0.5, y:3.25, w:6.0, h:1.0, color:'FFFFFF', fill:pptx.colors.ACCENT1 });
		slide.addText("ACCENT2", { x:7.0, y:3.25, w:6.0, h:1.0, color:'FFFFFF', fill:pptx.colors.ACCENT2 });
		slide.addText("ACCENT3", { x:0.5, y:4.70, w:6.0, h:1.0, color:'FFFFFF', fill:pptx.colors.ACCENT3 });
		slide.addText("ACCENT4", { x:7.0, y:4.70, w:6.0, h:1.0, color:'FFFFFF', fill:pptx.colors.ACCENT4 });
		slide.addText("ACCENT5", { x:0.5, y:6.15, w:6.0, h:1.0, color:'FFFFFF', fill:pptx.colors.ACCENT5 });
		slide.addText("ACCENT6", { x:7.0, y:6.15, w:6.0, h:1.0, color:'FFFFFF', fill:pptx.colors.ACCENT6 });

		// NEGATIVE TEST:
		//slide.addText("NEGTEST / NEGTEST", { x:0.5, y:0.5, w:'40%', h:0.38, color:pptx.colors.NEGTEST01, fill:pptx.colors.NEGTEST02 });
	}
}

function genSlides_Master(pptx) {
	var slide1 = pptx.addNewSlide('TITLE_SLIDE');
	slide1.addNotes('Master name: `TITLE_SLIDE`\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html');

	var slide2 = pptx.addNewSlide('MASTER_SLIDE');
	slide2.addNotes('Master name: `MASTER_SLIDE`\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html');
	slide2.addText('', { placeholder:'title' });

	var slide3 = pptx.addNewSlide('MASTER_SLIDE');
	slide3.addNotes('Master name: `MASTER_SLIDE` using pre-filled placeholders\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html');
	slide3.addText('Text Placeholder', { placeholder:'title' });
	slide3.addText(
		[
			{ text:'Pre-filled placeholder bullets', options:{ bullet:true, valign:'top' } },
			{ text:'Add any text, charts, whatever', options:{ bullet:true, indentLevel:1, color:'0000AB' } },
			{ text:'Check out the online API docs for more', options:{ bullet:true, indentLevel:2, color:'0000AB' } },
		],
		{ placeholder:'body', valign:'top' }
	);

	var slide4 = pptx.addNewSlide('MASTER_SLIDE');
	slide4.addNotes('Master name: `MASTER_SLIDE` using pre-filled placeholders\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html');
	slide4.addText('Image Placeholder', { placeholder:'title' });
	slide4.addImage({ placeholder:'body', path:(NODEJS ? gPaths.ccLogo.path.replace(/http.+\/examples/, '../examples') : gPaths.ccLogo.path) });

	var dataChartPieLocs = [
		{
			name  : 'Location',
			labels: ['CN', 'DE', 'GB', 'MX', 'JP', 'IN', 'US'],
			values: [  69,   35,   40,   85,   38,   99,  101]
		}
	];
	var slide5 = pptx.addNewSlide('MASTER_SLIDE');
	slide5.addNotes('Master name: `MASTER_SLIDE` using pre-filled placeholders\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html');
	slide5.addText('Chart Placeholder', { placeholder:'title' });
	slide5.addChart( pptx.charts.PIE, dataChartPieLocs, {showLegend:true, legendPos:'r', placeholder:'body'} );

	var slide6 = pptx.addNewSlide('THANKS_SLIDE');
	slide6.addNotes('Master name: `THANKS_SLIDE`\nAPI Docs: https://gitbrent.github.io/PptxGenJS/docs/masters.html');
	slide6.addText('Thank You!', { placeholder:'thanksText' });
	//slide6.addText('github.com/gitbrent', { placeholder:'body' });

	// LEGACY-TEST-ONLY: To check deprecated functionality
	/*
	if ( pptx.masters && Object.keys(pptx.masters).length > 0 ) {
		var slide1 = pptx.addNewSlide( pptx.masters.TITLE_SLIDE  );
		var slide2 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE );
		var slide3 = pptx.addNewSlide( pptx.masters.THANKS_SLIDE );

		var slide4 = pptx.addNewSlide( pptx.masters.TITLE_SLIDE,  { bkgd:'0088CC', slideNumber:{x:'50%', y:'90%', color:'0088CC'} } );
		var slide5 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:{ path:'https://raw.githubusercontent.com/gitbrent/PptxGenJS/v2.1.0/examples/images/title_bkgd_alt.jpg' } } );
		var slide6 = pptx.addNewSlide( pptx.masters.THANKS_SLIDE, { bkgd:'ffab33' } );
		//var slide7 = pptx.addNewSlide( pptx.masters.LEGACY_TEST_ONLY );
	}
	*/
}

// ==================================================================================================================

if ( typeof module !== 'undefined' && module.exports ) {
	module.exports = {
		execGenSlidesFuncs: execGenSlidesFuncs,
		runEveryTest: runEveryTest
	}
}
