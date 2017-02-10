/*
 * NAME: nodejs-demo.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: Jan 19, 2017
 * DESC: Demonstrate PptxGenJS on Node.js
 * REQS: npm 4.x + `npm install pptxgenjs`
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
var pptx = require('../dist/pptxgen.js'); // for LOCAL TESTING
//var pptx = require("pptxgenjs");
console.log(` * pptxgenjs version: ${pptx.getVersion()}`); // Loaded okay?

pptx.setTitle('PptxGenJS Node.js Demo');
pptx.setLayout('LAYOUT_WIDE');

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
var gStrHello = 'BONJOUR - CIAO - GUTEN TAG - HELLO - HOLA - NAMASTE - OLÀ - ZDRAS-TVUY-TE - 你好';
var optsTitle = { color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] };
var optsSubTitle = { x:0.5, y:0.7, cx:4, cy:0.3, font_size:18, font_face:'Arial', color:'0088CC', fill:'FFFFFF' };

// ============================================================================

function getTimestamp() {
	var dateNow = new Date();
	var dateMM = dateNow.getMonth() + 1; dateDD = dateNow.getDate(); dateYY = dateNow.getFullYear(), h = dateNow.getHours(); m = dateNow.getMinutes(); s = dateNow.getSeconds();
	return dateNow.getFullYear() +''+ (dateMM<=9 ? '0' + dateMM : dateMM) +''+ (dateDD<=9 ? '0' + dateDD : dateDD) + (h<=9 ? '0' + h : h) + (m<=9 ? '0' + m : m) + (s<=9 ? '0' + s : s);
}

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
	// ======== -----------------------------------------------------------------------------------
	// SLIDE 1: Misc Shape Types (no text)
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Shape Examples: Misc Shape Types (no text)', opts:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

	slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:8.00, y:0.75, cx:5, cy:3.25, fill:'00FF00' });
	slide.addShape(pptx.shapes.OVAL,              { x:4.15, y:0.75, cx:5, cy:2.00, fill:{ type:'solid', color:'0088CC', alpha:25 } });
	slide.addShape(pptx.shapes.RECTANGLE,         { x:0.50, y:0.75, cx:5, cy:3.25, fill:'FF0000' });
	slide.addShape(pptx.shapes.LINE,              { x:4.15, y:4.40, cx:5, cy:0, line:'FF0000', line_size:1 });
	slide.addShape(pptx.shapes.LINE,              { x:4.15, y:4.80, cx:5, cy:0, line:'FF0000', line_size:2, line_head:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.15, y:5.20, cx:5, cy:0, line:'FF0000', line_size:3, line_tail:'triangle' });
	slide.addShape(pptx.shapes.LINE,              { x:4.15, y:5.60, cx:5, cy:0, line:'FF0000', line_size:4, line_head:'triangle', line_tail:'triangle' });
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:0.40, y:4.30, cx:6, cy:3, fill:'0088CC', line:'000000', line_size:3 });
	slide.addShape(pptx.shapes.RIGHT_TRIANGLE,    { x:7.00, y:4.30, cx:6, cy:3, fill:'0088CC', line:'000000', flipH:true });

	// ======== -----------------------------------------------------------------------------------
	// SLIDE 2: Misc Shape Types with Text
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Shape Examples: Misc Shape Types (with text)', opts:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

	slide.addText('ROUNDED-RECTANGLE', { shape:pptx.shapes.ROUNDED_RECTANGLE, align:'l', x:8.00, y:0.75, cx:5, cy:3.25, fill:'00FF00' });
	slide.addText('OVAL',              { shape:pptx.shapes.OVAL,              align:'c', x:4.15, y:0.75, cx:5, cy:2.00, fill:{ type:'solid', color:'0088CC', alpha:25 } });
	slide.addText('RECTANGLE',         { shape:pptx.shapes.RECTANGLE,         align:'r', x:0.50, y:0.75, cx:5, cy:3.25, fill:'FF0000' });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'c', x:4.15, y:4.40, cx:5, cy:0, line:'FF0000', line_size:1 });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'l', x:4.15, y:4.80, cx:5, cy:0, line:'FF0000', line_size:2, line_head:'triangle' });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'r', x:4.15, y:5.20, cx:5, cy:0, line:'FF0000', line_size:3, line_tail:'triangle' });
	slide.addText('LINE',              { shape:pptx.shapes.LINE,              align:'c', x:4.15, y:5.60, cx:5, cy:0, line:'FF0000', line_size:4, line_head:'triangle', line_tail:'triangle' });
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE,    align:'c', x:0.40, y:4.30, cx:6, cy:3, fill:'0088CC', line:'000000', line_size:3 });
	slide.addText('RIGHT-TRIANGLE',    { shape:pptx.shapes.RIGHT_TRIANGLE,    align:'c', x:7.00, y:4.30, cx:6, cy:3, fill:'0088CC', line:'000000', flipH:true });
}

function demoImages() {
	// 1:
	var slide = pptx.addNewSlide();
	slide.addTable( [ [{ text:'Image Examples: Misc Image Types', opts:optsTitle }] ], { x:0.5, y:0.13, cx:12.5 } );

	// NODE.JS vvv: Load/encode Animated-GIF
	slide.addImage({ x:1.8, y:0.7, w:1.78, h:1.78, path:'images/anim_campfire.gif' });
	// NODE.JS ^^^:

	// Add an image using basic syntax
	slide.addImage({ path:'images/cc_copyremix.gif',          x:0.5, y:0.75, w:1.20, h:1.20 });
	// Slide API calls return the same slide, so you can chain calls:
	slide.addImage({ path:'images/cc_license_comp_chart.png', x:6.6, y:0.75, w:6.30, h:3.70 })
		 .addImage({ path:'images/cc_logo.jpg',               x:0.5, y:3.50, w:5.00, h:3.70 })
		 .addImage({ path:'images/cc_symbols_trans.png',      x:6.6, y:4.80, w:6.30, h:2.30 });

	// 2: Images can be pre-encoded into base64, so they do not have to be on the webserver etc. (saves generation time and resources!)
	// Also has the benefit of being able to be any type (path:images can only be exported as PNG)
	//slide.addImage({ x:1.8, y:0.7, w:1.78, h:1.78, data:GIF_ANIM_FIRE });
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
	[18,22,28,32].forEach(function(val,i){
		slide.addText( 'Misc font/color, size = '+val,
			{
				x:0.5, y:0.7+(i*0.7), cx:12.75,
				font_size:val, font_face:fonts[Math.floor((Math.random()*4)+1)], color:colors[i]
			}
		);
	});

	// Bullet Test: Number
	slide.addText(999, { x:9.0, y:0.6, w:'20%', h:1, color:'0000DE', bullet:true });
	// Bullet Test: Text test
	slide.addText('Bullet (str here, number above)', { x:9.0, y:1.0, w:'30%', h:1, color:'00AA00', bullet:true });

	// Multi-line test: bullet lines
	slide.addText('Line 1\nLine 2\nLine 3', { x:9.0, y:2.0, w:'30%', h:1, color:'393939', bullet:true, fill:'F2F2F2' });
	// Multi-line test: 3 lines
	slide.addText(
		'***Line-Break/Multi-Line Test***\n\nFirst line\nSecond line\nThird line',
		{ x:0.4, y:4.0, w:5, h:2, valign:'middle', align:'ctr', color:'6c6c6c', font_size:16, fill:'F2F2F2' }
	);

	// Effects > Shadow
	var shadowOpts = { type:'outer', color:'696969', blur:3, offset:10, angle:45, opacity:0.8 };
	slide.addText('Outer Shadow (blur:3, offset:10, angle:45, opacity:80%)', { x:0.5, y:6.0, w:12, h:1, font_size:32, color:'0088cc', shadow:shadowOpts });

	// ======== -----------------------------------------------------------------------------------
	// SLIDE 2: Misc mess
	// ======== -----------------------------------------------------------------------------------
	var slide = pptx.addNewSlide();
	// Slide colors: bkgd/fore
	slide.back = '030303';
	slide.color = '9F9F9F';
	// Title
	slide.addTable( [ [{ text:'Text Examples 2', opts:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

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

function demoTables() {
	// SLIDE 1: Table text alignment and cell styles
	// ======== -----------------------------------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		slide.addTable( [ [{ text:'Table Examples 1', opts:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

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
			arrTabRows, { x:0.5, y:1.1, w:6.0 },
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
			arrTabRows, { x:7.0, y:1.1, w:6.0 },
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
	// ======== -----------------------------------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		// 2: Slide title
		slide.addTable(
			[ [{ text:'Table Examples 2', opts:{ color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] } }] ],
			{ x:0.5, y:0.13, cx:12.5 }
		);

		// DEMO: Row/Col Width/Heights ----------------------------------------------------------------
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle));
		slide.addText('Colspans/Rowspans:', optsSub);

		var tabOpts1 = { x:0.5, y:1.1, w:12.0, rowH:1.0, fill:'F5F5F5', color:'3D3D3D', font_size:16, border:'FFFFFF', align:'c', valign:'c' };
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
		slide.addTable( arrTabRows1, tabOpts1 );

		var arrTabRows2 = [
			[ { text:'A1/B1', opts:{colspan:2, fill:'00CC00'} }, { text:'C1' } ],
			[ { text:'A2' }, { text:'B2/C2', opts:{colspan:2, fill:'00cc00'} } ]
		];
		// NOTE: Follow HTML conventions for colspan/rowspan cells - cells spanned are left out of arrays - see above
		// The table above has 6 columns, but each of the 3 rows has 4-5 elements as colspan/rowspan replacing the missing ones
		// (e.g.: there are 5 elements in the first row, and 6 in the second)
		var tabOpts2 = { x:0.5, y:4.8, w:12.0, rowH:1.0, fill:'F5F5F5', color:'3D3D3D', font_size:16, border:'FFFFFF', align:'c', valign:'c' };
		slide.addTable( arrTabRows2, tabOpts2 );
	}

	// SLIDE 3: Cell Formatting / Cell Margins
	// ======== -----------------------------------------------------------------------------------
	{
		var slide = pptx.addNewSlide();
		// 2: Slide title
		slide.addTable(
			[ [{ text:'Table Examples 3', opts:{ color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] } }] ],
			{ x:0.5, y:0.13, cx:12.5 }
		);

		// Cell Margins
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle));
		slide.addText('Cell Margins:', optsSub);

		slide.addTable( [['margin:0']],           { x:0.5, y:1.1, margin:0,           w:1.2, fill:'FFFCCC' } );
		slide.addTable( [['margin:[0,0,0,20']],   { x:0.5, y:1.6, margin:[0,0,0,20],  w:1.2, fill:'FFFCCC' } );
		slide.addTable( [['margin:5']],           { x:2.5, y:1.1, margin:5,           w:1.0, fill:'F1F1F1' } );
		slide.addTable( [['margin:[40,5,5,20]']], { x:4.5, y:1.1, margin:[40,5,5,20], w:2.2, fill:'F1F1F1' } );
		slide.addTable( [['margin:[80,5,5,10]']], { x:8.0, y:1.1, margin:[80,5,5,10], w:2.2, fill:'F1F1F1' } );

		// Complex/Compound border
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle)); optsSub.y = 3.0;
		slide.addText('Complex Cell Border:', optsSub);
		var arrBorder = [ {color:'FF0000',pt:1}, {color:'00ff00',pt:3}, {color:'0000ff',pt:5}, {color:'9e9e9e',pt:7} ];
		slide.addTable( [['Borders!']], { x:0.5, y:3.4, w:12.3, rowH:2.0, fill:'F5F5F5', color:'3D3D3D', font_size:18, border:arrBorder, align:'c', valign:'c' } );

		// Invalid char check
		var optsSub = JSON.parse(JSON.stringify(optsSubTitle)); optsSub.y = 6.1;
		slide.addText('Escaped Invalid Chars:', optsSub);
		var arrTabRows3 = [['<', '>', '"', "'", '&', 'plain']];
		slide.addTable( arrTabRows3, { x:0.5, y:6.5, w:12.3, rowH:0.5, fill:'F5F5F5', color:'3D3D3D', border:'FFFFFF', align:'c', valign:'c' } );

	}

	// SLIDE 4: Table auto-paging
	// ======== -----------------------------------------------------------------------------------
	{
		var arrRows = [];
		for (var idx=0; idx<gArrNamesF.length; idx++) {
			arrRows.push( [idx, gArrNamesF[idx], gStrLorumIpsum.substring(idx*100,idx*300)] );
		}

		pptx.addNewSlide().addTable( arrRows, { x:0.5, y:0.25, colW:[0.75,1.75,10], margin:2, border:'CFCFCF' } );

		pptx.addNewSlide().addTable( arrRows, { x:3.0, y:0.25, colW:[0.75,1.75, 7], margin:5, border:'CFCFCF' } );
	}
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

function demoMedia() {
	var optsTitle = { color:'9F9F9F', marginPt:3, border:[0,0,{pt:'1',color:'CFCFCF'},0] };

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
	slide1.addMedia({ x:9.4, y:4.0, w:3.00, h:2.25, type:'online', link:'https://www.youtube.com/embed/Dph6ynRVyUc' });

	var slide2 = pptx.addNewSlide();
	slide2.addTable( [ [{ text:'Media: Audio Examples', opts:optsTitle }] ], { x:0.5, y:0.13, w:12.5 } );

	slide2.addText('Audio: mp3', { x:0.5, y:0.6, w:4.00, h:0.4, color:'0088CC' });
	slide2.addMedia({ x:0.5, y:1.0, w:4.00, h:0.3, type:'audio', path:'media/sample.mp3' });
	slide2.addMedia({ x:0.5, y:3.0, w:4.00, h:0.3, type:'audio', path:'media/sample.wav' });
}

// ============================================================================

// STEP 3: Run all test funcs
demoBasic();
demoText();
demoImages();
demoShapes();
demoTables();
demoMedia();
demoCorpPres();

// STEP 4: Export giant demo file
pptx.save('Node_Demo_'+getTimestamp());

// ============================================================================

console.log(`
--------------
TEST COMPLETE!
--------------
 * Files saved to...: ${__dirname}
`);
