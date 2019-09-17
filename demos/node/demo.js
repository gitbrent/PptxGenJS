/*
 * NAME: demo.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: 20190916
 * DESC: PptxGenJS feature demos for Node.js
 * REQS: npm 4.x + `npm install pptxgenjs`
 *
 * USAGE: `node demo.js`       (runs local tests with callbacks etc)
 * USAGE: `node demo.js All`   (runs all pre-defined tests in `../common/demos.js`)
 * USAGE: `node demo.js Text`  (runs pre-defined single test in `../common/demos.js`)
 */

// ============================================================================
const express = require('express'); // Not core - Only required for streaming
const app = express(); // Not core - Only required for streaming
let verboseMode = true;
let PptxGenJS;

function getTimestamp() {
	var dateNow = new Date();
	var dateMM = dateNow.getMonth() + 1; dateDD = dateNow.getDate(); dateYY = dateNow.getFullYear(), h = dateNow.getHours(); m = dateNow.getMinutes();
	return dateNow.getFullYear() +''+ (dateMM<=9 ? '0' + dateMM : dateMM) +''+ (dateDD<=9 ? '0' + dateDD : dateDD) + (h<=9 ? '0' + h : h) + (m<=9 ? '0' + m : m);
}
// ============================================================================

if (verboseMode) console.log(`
-------------
STARTING DEMO
-------------
`);

// STEP 1: Load pptxgenjs library
if ( (process.argv[2] && process.argv[2].toLowerCase() == '-local') || (process.argv[3] && process.argv[3].toLowerCase() == '-local' ) ) {
	// for LOCAL TESTING
	PptxGenJS = require('../../dist/pptxgen.cjs.js');
	if (verboseMode) console.log('--=== LOCAL MODE ===--');
	let pptx = new PptxGenJS();
	if (verboseMode) console.log(`* pptxgenjs ver: ${pptx.version}`);
}
else {
	PptxGenJS = require("pptxgenjs");
}
var pptx = new PptxGenJS();
var demo = require("../common/demos.js");

if (verboseMode) console.log(`* save location: ${__dirname}`);

// ============================================================================

// EX: Regular callback - will be sent the export filename once the file has been written to fs
function saveCallback(filename) {
	if (verboseMode) {
		console.log('`saveCallback()`: Export filename: '+ filename);
	}
}

// EX: JSZip callback - take the specified output (`data`) and do whatever
function jszipCallback(data) {
	if (verboseMode) {
		console.log('jszipCallback(): Here are 0-100 chars of `data`:\n');
		console.log( data.substring(0,100) );
	}
}

// EX: Callback that receives the PPT binary data - use this to stream file
function streamCallback(data) {
	var strFilename = "Node-Demo-Streamed-Callback.pptx";

	app.get('/', (req,res) => {
		res.writeHead(200, { 'Content-disposition':'attachment;filename='+strFilename, 'Content-Length':data.length });
		res.end(new Buffer(data, 'binary'));
	});

	app.listen(3000, () => {
		console.log('PptxGenJS Node Demo app listening on port 3000!');
		console.log('Visit: http://localhost:3000/');
		console.log('(press Ctrl-C to quit app)');
	});
}

// ============================================================================

// STEP 2: Run predefined test from `../common/demos.js` //-OR-// Local Tests (callbacks, etc.)
if ( process.argv.length == 3 ) {
	if ( process.argv[2].toLowerCase() == 'all' ) demo.runEveryTest();
	else demo.execGenSlidesFuncs( process.argv[2] );
}
else {
	// STEP 3: Omit an arg to run only these below
	var exportName = 'PptxGenJS_Demo_Node_'+getTimestamp();
	var pptx = new PptxGenJS();
	var slide = pptx.addSlide();
	slide.addText( 'New Node Presentation', {x:1.5, y:1.5, w:6, h:2, margin:0.1, fill:'FFFCCC'} );
	// Test that `pptxgen.shapes.js` was loaded/is available
	slide.addShape(pptx.shapes.OVAL_CALLOUT, { x:6, y:2, w:3, h:2, fill:'00FF00', line:'000000', lineSize:1 });

	// **NOTE**: Only uncomment one EXAMPLE at a time

	// EXAMPLE 1: Inline save (saves to the local directory where this process is running)
	//pptx.save( exportName+'-ex1' ); if (verboseMode) console.log('\nFile created:\n'+' * '+exportName+'-ex1');

	// EXAMPLE 2: Use an inline callback function
	pptx.writeFile(exportName+'-ex2')
		.catch(ex => { console.log('ERROR: '+err) })
		.then(fileName => { console.log('Ex2 inline callback exported: '+fileName); } );

	// EXAMPLE 3: Use defined callback function
	//pptx.save( exportName+'-ex3', saveCallback );

	// EXAMPLE 4: Save in various formats - JSZip offers: ['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array']
	//pptx.save( exportName+'-jszip', jszipCallback, 'base64' );

	// EXAMPLE 5: Use callback with 'http' in filename to get content back instead of writing a file - use this for streaming, cloud, etc.
	// TEST: Use your browser and go to "http://localhost:3000/", then the file will be downloaded
	//pptx.save( 'http', streamCallback );

	// **NOTE** If you continue to use the `pptx` variable, new Slides will be added to the existing set
}

// ============================================================================

if (verboseMode) console.log(`
--------------
DEMO COMPLETE!
--------------
`);
