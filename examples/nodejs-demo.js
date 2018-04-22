/*
 * NAME: nodejs-demo.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: 20180421
 * DESC: Demonstrate PptxGenJS on Node.js
 * REQS: npm 4.x + `npm install pptxgenjs`
 * EXEC: `node nodejs-demo.js`
 * EXEC: `node nodejs-demo.js All`
 * EXEC: `node nodejs-demo.js Text`
 */

// ============================================================================
const express = require('express'); // Not core - Only required for streaming
const app = express(); // Not core - Only required for streaming
const fs = require('fs');
var gConsoleLog = true;

function getTimestamp() {
	var dateNow = new Date();
	var dateMM = dateNow.getMonth() + 1; dateDD = dateNow.getDate(); dateYY = dateNow.getFullYear(), h = dateNow.getHours(); m = dateNow.getMinutes();
	return dateNow.getFullYear() +''+ (dateMM<=9 ? '0' + dateMM : dateMM) +''+ (dateDD<=9 ? '0' + dateDD : dateDD) + (h<=9 ? '0' + h : h) + (m<=9 ? '0' + m : m);
}
// ============================================================================

if (gConsoleLog) console.log(`
-------------
STARTING DEMO
-------------
`);

// STEP 1: Load pptxgenjs and show version to verify everything loaded correctly
var PptxGenJS;
if (fs.existsSync('../dist/pptxgen.js')) {
	// for LOCAL TESTING
	PptxGenJS = require('../dist/pptxgen.js');
	if (gConsoleLog) console.log('-=TEST MODE=- (../dist/pptxgen.js)');
}
else {
	PptxGenJS = require("pptxgenjs");
}
var pptx = new PptxGenJS();
var demo = require("../examples/pptxgenjs-demo.js");

if (gConsoleLog) console.log(` * save location: ${__dirname}`);

// ============================================================================

// EX: Regular callback - will be sent the export filename once the file has been written to fs
function saveCallback(filename) {
	if (gConsoleLog) console.log('saveCallback: Good News Everyone! File created: '+ filename);
}

// EX: JSZip callback - take the specified output (`data`) and do whatever
function jszipCallback(data) {
	if (gConsoleLog) {
		console.log('jszipCallback: First 100 chars of output:\n');
		console.log( data.substring(0,100) );
	}
}

// EX: Callback that receives the PPT binary data - use this to stream file
function streamCallback(data) {
	var strFilename = "Node-Presenation-Streamed.pptx";

	app.get('/', function(req, res) {
		res.writeHead(200, { 'Content-disposition':'attachment;filename='+strFilename, 'Content-Length':data.length });
		res.end(new Buffer(data, 'binary'));
	});

	app.listen(3000, function() {
		console.log('PptxGenJS Node Demo app listening on port 3000!');
		console.log('Visit: http://localhost:3000/');
		console.log('(press Ctrl-C to quit app)');
	});
}

// ============================================================================

// STEP 2: Run specified test, or all test funcs
if ( process.argv.length == 3 ) {
	if ( process.argv[2] == 'All' || process.argv[2] == 'all' ) demo.runEveryTest();
	else demo.execGenSlidesFuncs(process.argv[2]);
}

// STEP 3: Export another demo file
// HOWTO: Create a new Presenation
var pptx = new PptxGenJS();
if (gConsoleLog && process.argv.length != 3) console.log(` * pptxgenjs ver: ${pptx.version}`); // Loaded okay?

var exportName = 'PptxGenJS_Demo_Node2_'+getTimestamp();
var slide = pptx.addNewSlide();
slide.addText( 'New Node Presentation', {x:1.5, y:1.5, w:6, h:2, margin:0.1, fill:'FFFCCC'} );

// A: Inline save
pptx.save( exportName ); if (gConsoleLog) console.log('\nFile created:\n'+' * '+exportName);

// B: or Save using callback function
//pptx.save( exportName, function(filename){ console.log('Inline callback here! -> '+exportName); } ); if (gConsoleLog) console.log('\nFile created:\n'+' * '+exportName);

// C: or use a predefined callback function
//pptx.save( exportName, saveCallback ); if (gConsoleLog) console.log('\nFile created:\n'+' * '+exportName);

// D: or use callback with 'http' in filename to get content back instead of writing a file - use this for streaming
//pptx.save( 'http', streamCallback );

// E: or save using various JSZip formats ['arraybuffer', 'base64', 'binarystring', 'blob', 'nodebuffer', 'uint8array']
//pptx.save( 'jszip', jszipCallback, 'base64' );

// **NOTE** If you continue to use the `pptx` variable, new Slides will be added to the existing set

// ============================================================================

if (gConsoleLog) console.log(`
--------------
DEMO COMPLETE!
--------------
`);
