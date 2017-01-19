/*
 * NAME: nodejs-demo.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: Jan 19, 2017
 * DESC: Demonstrate PptxGenJS on Node.js
 * REQS: Node 4.x + `npm install pptxgenjs`
 * EXEC: `node nodejs-demo.js`
 */

 const colors = ['FF0000','AB00CD','00FF00','00AA00','003300','330033','990099','33FFFF','AA33CC','336699'];
 const fonts = ['Arial','Courier New','Times','Verdana'];

console.log(`
-------------
STARTING TEST
-------------`);

// ============================================================================

// STEP 1: Load pptxgenjs and show version to verify everything loaded correctly
//var pptx = require('../dist/pptxgen.js'); // for LOCAL TESTING
var pptx = require("pptxgenjs");
console.log(` * pptxgenjs version: ${pptx.getVersion()}`); // Loaded okay?

pptx.setTitle('PptxGenJS Node.js Demo');
pptx.setLayout('LAYOUT_WIDE');

var optsTitle = { color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] };
var optsSubTitle = { x:0.5, y:0.7, cx:4, cy:0.3, font_size:18, font_face:'Arial', color:'0088CC', fill:'FFFFFF' };

// ============================================================================

// STEP 2: Define Demo funcs
function demoBasic() {
	var slide = pptx.addNewSlide();
	var optsTitle = { color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] };

	slide.slideNumber({ x:0.5, y:'90%' });

	slide.addTable( [ [{ text:'Simple Example', opts:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );
	slide.addText('Hello World!', { x:0.5, y:0.7, w:6, h:1, color:'0000FF' });
	// Bullet Test: Number
	slide.addText(999, { x:0.5, y:2.0, w:'50%', h:1, color:'0000DE', bullet:true });
	// Bullet Test: Text test
	slide.addText('Bullet text', { x:0.5, y:2.5, w:'50%', h:1, color:'00AA00', bullet:true });
	// Bullet Test: Multi-line text test
	slide.addText('Line 1\nLine 2\nLine 3', { x:0.5, y:3.5, w:'50%', h:1, color:'AACD00', bullet:true });
}

function demoShapes() {
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Shape Examples: Misc Shape Types (no text)', opts:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

	slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:8.00, y:0.75, cx:5, cy:3.25, fill:'00FF00' });
	slide.addShape(pptx.shapes.OVAL,              { x:4.15, y:0.75, cx:5, cy:2.00, fill:{ type:'solid', color:'0088CC', alpha:25 } });
	slide.addShape(pptx.shapes.RECTANGLE,         { x:0.50, y:0.75, cx:5, cy:3.25, fill:'FF0000' });
	slide.addShape(pptx.shapes.LINE,              { x:4.15, y:4.4, cx:5, cy:0, line:'FF0000', line_size:1 });
	slide.addShape(pptx.shapes.LINE,              { x:4.15, y:4.8, cx:5, cy:0, line:'FF0000', line_size:2, line_head:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.15, y:5.2, cx:5, cy:0, line:'FF0000', line_size:3, line_tail:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.15, y:5.6, cx:5, cy:0, line:'FF0000', line_size:4, line_head:'triangle', line_tail:'triangle' });
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:0.40, y:4.3, cx:6, cy:3, fill:'0088CC', line:'000000', line_size:3 });
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:7.00, y:4.3, cx:6, cy:3, fill:'0088CC', line:'000000', flipH:true });

	// ======== -----------------------------------------------------------------------------------
	// SLIDE 2: Misc Shape Types with Text
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Shape Examples: Misc Shape Types (with text)', opts:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

	slide.addText('ROUNDED-RECTANGLE', { shape:pptx.shapes.ROUNDED_RECTANGLE, align:'l', x:8.00, y:0.75, cx:5, cy:3.25, fill:'00FF00' });
	slide.addText('OVAL',              { shape:pptx.shapes.OVAL,              align:'c', x:4.15, y:0.75, cx:5, cy:2.00, fill:{ type:'solid', color:'0088CC', alpha:25 } });
	slide.addText('RECTANGLE',         { shape:pptx.shapes.RECTANGLE,         align:'r', x:0.50, y:0.75, cx:5, cy:3.25, fill:'FF0000' });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'c', x:4.15, y:4.40, cx:5, cy:0, line:'FF0000', line_size:1 });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'l', x:4.15, y:4.8, cx:5, cy:0, line:'FF0000', line_size:2, line_head:'triangle' });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'r', x:4.15, y:5.2, cx:5, cy:0, line:'FF0000', line_size:3, line_tail:'triangle' });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'c', x:4.15, y:5.6, cx:5, cy:0, line:'FF0000', line_size:4, line_head:'triangle', line_tail:'triangle' });
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE,    align:'c', x:0.40, y:4.3, cx:6, cy:3, fill:'0088CC', line:'000000', line_size:3 });
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE,    align:'c', x:7.00, y:4.3, cx:6, cy:3, fill:'0088CC', line:'000000', flipH:true });
}

function demoImages() {
	// 1:
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Image Examples: Misc Image Types', opts:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

	// Add an image using basic syntax
	slide.addImage({ path:'images/cc_copyremix.gif',          x:0.5, y:0.75 }); // TEST for image without size provided

	// Slide API calls return the same slide, so you can chain calls:
	slide.addImage({ path:'images/cc_license_comp_chart.png', x:6.6, y:0.75, w:6.30, h:3.70 })
		 .addImage({ path:'images/cc_logo.jpg',               x:0.5, y:3.50, w:5.00, h:3.70 })
		 .addImage({ path:'images/cc_symbols_trans.png',      x:6.6, y:4.80, w:6.30, h:2.30 });

	// Node.js TEST: Load/encode Animated-GIF
	slide.addImage({ x:1.8, y:0.7, w:1.78, h:1.78, path:'images/anim_campfire.gif' });

	// NOTE: The 'data:' part of the encoded string is optional:
	slide.addImage({ x:3.7, y:1.3, w:0.6, h:0.6, data:'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAAjcAAAI3AGf6F88AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAANVQTFRF////JLaSIJ+AIKqKKa2FKLCIJq+IJa6HJa6JJa6IJa6IJa2IJa6IJa6IJa6IJa6IJa6IJa6IJq6IKK+JKK+KKrCLLrGNL7KOMrOPNrSRN7WSPLeVQrmYRLmZSrycTr2eUb6gUb+gWsKlY8Wqbsmwb8mwdcy0d8y1e863g9G7hdK8htK9i9TAjNTAjtXBktfEntvKoNzLquDRruHTtePWt+TYv+fcx+rhyOvh0e7m1e/o2fHq4PTu5PXx5vbx7Pj18fr49fv59/z7+Pz7+f38/P79/f7+dNHCUgAAABF0Uk5TAAcIGBktSYSXmMHI2uPy8/XVqDFbAAABB0lEQVQ4y42T13qDMAyFZUKMbebp3mmbrnTvlY60TXn/R+oFGAyYzz1Xx/wylmWJqBLjUkVpGinJGXXliwSVEuG3sBdkaCgLPJMPQnQUDmo+jGFRPKz2WzkQl//wQvQoLPII0KuAiMjP+gMyn4iEFU1eAQCCiCU2fpCfFBVjxG18f35VOk7Swndmt9pKUl2++fG4qL2iqMPXpi8r1SKitDDne/rT8vPbRh2d6oC7n6PCLNx/bsEM0Edc5DdLAHD9tWueF9VJjmdP68DZ77iRkDKuuT19Hx3mx82MpVmo1Yfv+WXrSrxZ6slpiyes77FKif88t7Nh3C3nbFp327sHxz167uHtH/8/eds7gGsUQbkAAAAASUVORK5CYII=' });
	// TEST: Ensure framework corrects for missing type header
	slide.addImage({ x:4.4, y:1.9, w:0.7, h:0.7, data:'base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAAAA3NCSVQICAjb4U/gAAAACXBIWXMAAAjcAAAI3AGf6F88AAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAANVQTFRF////JLaSIJ+AIKqKKa2FKLCIJq+IJa6HJa6JJa6IJa6IJa2IJa6IJa6IJa6IJa6IJa6IJa6IJq6IKK+JKK+KKrCLLrGNL7KOMrOPNrSRN7WSPLeVQrmYRLmZSrycTr2eUb6gUb+gWsKlY8Wqbsmwb8mwdcy0d8y1e863g9G7hdK8htK9i9TAjNTAjtXBktfEntvKoNzLquDRruHTtePWt+TYv+fcx+rhyOvh0e7m1e/o2fHq4PTu5PXx5vbx7Pj18fr49fv59/z7+Pz7+f38/P79/f7+dNHCUgAAABF0Uk5TAAcIGBktSYSXmMHI2uPy8/XVqDFbAAABB0lEQVQ4y42T13qDMAyFZUKMbebp3mmbrnTvlY60TXn/R+oFGAyYzz1Xx/wylmWJqBLjUkVpGinJGXXliwSVEuG3sBdkaCgLPJMPQnQUDmo+jGFRPKz2WzkQl//wQvQoLPII0KuAiMjP+gMyn4iEFU1eAQCCiCU2fpCfFBVjxG18f35VOk7Swndmt9pKUl2++fG4qL2iqMPXpi8r1SKitDDne/rT8vPbRh2d6oC7n6PCLNx/bsEM0Edc5DdLAHD9tWueF9VJjmdP68DZ77iRkDKuuT19Hx3mx82MpVmo1Yfv+WXrSrxZ6slpiyes77FKif88t7Nh3C3nbFp327sHxz167uHtH/8/eds7gGsUQbkAAAAASUVORK5CYII=' });
}

function demoText() {
	// ======== -----------------------------------------------------------------------------------
	// SLIDE 1: Font Size/Color line examples 12-60 pt font, random color
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Text Examples 1', opts:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

	// DEMO: Array of text fonts/colors in increasing size
	[12,18,22,28,32,42,48,60].forEach(function(val,i){
		slide.addText( 'Misc font/color, size = '+val,
			{
				x:0.5, y:0.7+(i*0.8), cx:12.75,
				font_size:val, font_face:fonts[Math.floor((Math.random()*4)+1)], color:colors[i]
			}
		);
	});

	// Bullet Test: Number
	slide.addText(999, { x:9.0, y:0.8, w:'20%', h:1, color:'0000DE', bullet:true });
	// Bullet Test: Text test
	slide.addText('Bullet text (int above)', { x:9.0, y:1.5, w:'30%', h:1, color:'00AA00', bullet:true });
	// Bullet Test: Multi-line text test
	slide.addText('Line 1\nLine 2\nLine 3', { x:9.0, y:2.5, w:'30%', h:1, color:'AACD00', bullet:true });

	// ======== -----------------------------------------------------------------------------------
	// SLIDE 2: Misc mess
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	// Slide colors: bkgd/fore
	slide.back = '030303';
	slide.color = '9F9F9F';
	// Title
	slide.addTable( [ [{ text:'Text Examples 2', opts:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

	// Actual Textbox shape (can have any Height, can wrap text, etc.)
	slide.addText( 'Textbox (ctr/ctr)', { x:0.5, y:0.75, cx:8.5, cy:2.5, color:'FFFFFF', fill:'0000FF', valign:'c', align:'c', isTextBox:true } );
	slide.addText( 'Textbox (top/lft)', { x:10,  y:0.75, cx:3.0, cy:1.0, color:'FFFFFF', fill:'00CC00', valign:'t', align:'l', isTextBox:true } );
	slide.addText( 'Textbox (btm/rgt)', { x:10,  y:2.25, cx:3.0, cy:1.0, color:'FFFFFF', fill:'FF0000', valign:'b', align:'r', isTextBox:true } );

	slide.addText('Plain x/y coords', { x:10, y:3.5 });

	slide.addText('Escaped chars: \' " & < >', { x:10, y:4.5 });

	slide.addText('^ (50%/50%)', {x:'50%', y:'50%'});

	// Add text box with multi colors and fonts:
	slide.addText(
		[
			{ text:'Hello ', options:{ font_size:48, font_face:'Courier New', color:'0000CC' } },
			{ text:'World!', options:{ font_size:48, font_face:'Arial', color:'FFFF00' } }
		],
		{ x:0.5, y:2.5, cx:2.5, cy:3, margin:0.25 }
	);

	var objOptions = {
		x:0.5, y:5.5, cx:'90%',
		font_face:'Arial', font_size:42, color:'00CC00', bold:true, italic:true, underline:true, margin:0, isTextBox:true
	};
	slide.addText('Arial (42pt, green, bold, italic, underline), [0 inset]', objOptions);

	slide.addText('Footer Bar: PptxGenJS version ' + pptx.getVersion() + ' (width:100%, valign:ctr)',
		{ x:0, y:6.5, cx:'100%', cy:0.75, fill:'f7f7f7', color:'666666', align:'center', valign:'middle' }
	);
}

function demoTables() {
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Table Examples 1', opts:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

	// DEMO: align/valign -------------------------------------------------------------------------
	var objOpts1 = { x:0.5, y:0.7, font_size:18, font_face:'Arial', color:'0088CC' };
	slide.addText('Cell Text Alignment:', objOpts1);

	var arrTabRows = [
		[
			{ text: 'Top Lft', opts: { valign:'top',    align:'left'  , font_face:'Arial'   } },
			{ text: 'Top Ctr', opts: { valign:'t'  ,    align:'center', font_face:'Courier' } },
			{ text: 'Top Rgt', opts: { valign:'t'  ,    align:'right' , font_face:'Verdana' } }
		],
		[
			{ text: 'Ctr Lft', opts: { valign:'middle', align:'left' } },
			{ text: 'Ctr Ctr', opts: { valign:'center', align:'ctr'  } },
			{ text: 'Ctr Rgt', opts: { valign:'c'     , align:'r'    } }
		],
		[
			{ text: 'Btm Lft', opts: { valign:'bottom', align:'l' } },
			{ text: 'Btm Ctr', opts: { valign:'btm',    align:'c' } },
			{ text: 'Btm Rgt', opts: { valign:'b',      align:'r' } }
		]
	];
	slide.addTable(
		arrTabRows, { x:0.5, y:1.1, cx:6.0 },
		{ rowH:0.75, fill:'F7F7F7', font_size:14, color:'363636', border:{pt:'1', color:'BBCCDD'} }
	);
	// Pass default cell style as tabOpts, then just style/override individual cells as needed

	// DEMO: cell styles --------------------------------------------------------------------------
	var objOpts2 = { x:7.0, y:0.7, font_size:18, font_face:'Arial', color:'0088CC' };
	slide.addText('Cell Styles:', objOpts2);

	var arrTabRows = [
		[
			{ text: 'White',  opts: { fill:'6699CC', color:'FFFFFF' } },
			{ text: 'Yellow', opts: { fill:'99AACC', color:'FFFFAA' } },
			{ text: 'Pink',   opts: { fill:'AACCFF', color:'E140FE' } }
		],
		[
			{ text: '12pt', opts: { fill:'FF0000', font_size:12 } },
			{ text: '20pt', opts: { fill:'00FF00', font_size:20 } },
			{ text: '28pt', opts: { fill:'0000FF', font_size:28 } }
		],
		[
			{ text: 'Bold',      opts: { fill:'003366', bold:true } },
			{ text: 'Underline', opts: { fill:'336699', underline:true } },
			{ text: '10pt Pad',  opts: { fill:'6699CC', marginPt:10 } }
		]
	];
	slide.addTable(
		arrTabRows, { x:7.0, y:1.1, cx:6.0 },
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

	// ======== -----------------------------------------------------------------------------------
	// SLIDE 2: Table row/col-spans
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	// 2: Slide title
	slide.addTable(
		[ [{ text:'Table Examples 2', opts:{ color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] } }] ],
		{ x:0.5, y:0.13, cx:12.5 }
	);

	// DEMO: Row/Col Width/Heights ----------------------------------------------------------------
	var optsSub = JSON.parse(JSON.stringify(optsSubTitle));
	slide.addText('Row/Col-spans:', optsSub);

	var arrTabRows1 = [
		[
			{ text:'A1 and A2', opts:{rowspan:2, fill:'00CC00'} }
			,{ text:'B1' }
			,{ text:'C1 and D1', opts:{colspan:2, fill:'00CC00'} }
			,{ text:'E1' }
			,{ text:'F1/F2/F3',  opts:{rowspan:3, fill:'00CC00'} }
		]
		,[       'B2', 'C2', 'D2', 'E2' ]
		,[ 'A3', 'B3', 'C3', 'D3', 'E3' ]
	];
	// NOTE: Follow HTML conventions for colspan/rowspan cells - cells spanned are left out of arrays - see above
	// The table above has 6 columns, but each of the 3 rows has 4-5 elements as colspan/rowspan replacing the missing ones
	// (e.g.: there are 5 elements in the first row, and 6 in the second)
	slide.addTable( arrTabRows1, { x:0.5, y:1.1, w:12.3, rowH:0.75, fill:'F5F5F5', color:'3D3D3D', border:'FFFFFF', align:'c', valign:'c' } );

	var arrTabRows2 = [
		[ { text:'A1/B1', opts:{colspan:2, fill:'00CC00'} }, { text:'C1' } ],
		[ { text:'A2' }, { text:'B2/C2', opts:{colspan:2, fill:'00cc00'} } ]
	];
	// NOTE: Follow HTML conventions for colspan/rowspan cells - cells spanned are left out of arrays - see above
	// The table above has 6 columns, but each of the 3 rows has 4-5 elements as colspan/rowspan replacing the missing ones
	// (e.g.: there are 5 elements in the first row, and 6 in the second)
	slide.addTable( arrTabRows2, { x:0.5, y:3.6, w:12.3, rowH:0.5, fill:'F5F5F5', color:'3D3D3D', border:'FFFFFF', align:'c', valign:'c' } );

	// Complex/Compound border
	var optsSub = JSON.parse(JSON.stringify(optsSubTitle)); optsSub.y = 4.8;
	slide.addText('Complex Cell Border:', optsSub);
	var arrBorder = [ {color:'FF0000',pt:1}, {color:'00ff00',pt:3}, {color:'0000ff',pt:5}, {color:'9e9e9e',pt:7} ];
	slide.addTable( [['Borders!']], { x:0.5, y:5.2, w:12.3, rowH:0.6, fill:'F5F5F5', color:'3D3D3D', border:arrBorder, align:'c', valign:'c' } );

	// Invalid char check
	var optsSub = JSON.parse(JSON.stringify(optsSubTitle)); optsSub.y = 6.1;
	slide.addText('Escaped Invalid Chars:', optsSub);
	var arrTabRows3 = [['<', '>', '"', "'", '&', 'plain']];
	slide.addTable( arrTabRows3, { x:0.5, y:6.5, w:12.3, rowH:0.5, fill:'F5F5F5', color:'3D3D3D', border:'FFFFFF', align:'c', valign:'c' } );
}

function demoCorpPres() {
	//pptx.setLayout('LAYOUT_16x9'); // NOTE: Master slide demos are 16x9, so they have lots of whitepsace in this demo...

	var slide1 = pptx.addNewSlide( pptx.masters.TITLE_SLIDE  );
	var slide2 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE );
	var slide3 = pptx.addNewSlide( pptx.masters.THANKS_SLIDE );

	var slide4 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:'0088CC'} );
	// FIXME: TODO: this image doesnt enocde right (the rels.data doesnt start with "data:image/png;base64,iV..", its "/9j/4..")
	//var slide5 = pptx.addNewSlide( pptx.masters.MASTER_SLIDE, { bkgd:{ src:'images/title_bkgd_alt.jpg' } } ); // TEST: override bkgd
	var slide6 = pptx.addNewSlide( pptx.masters.THANKS_SLIDE, { bkgd:'ffab33'} );
}

// ============================================================================

// STEP 3: Run all test funcs
demoBasic();
demoText();
demoImages();
demoShapes();
demoTables();
demoCorpPres();

// STEP 4: Export giant demo file
pptx.save('Nodejs-Demo-Presentation');

// ============================================================================

console.log(`
--------------
TEST COMPLETE!
--------------
 * Files saved to...: ${__dirname}
`);
