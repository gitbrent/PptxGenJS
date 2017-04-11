/*
 * NAME: nodejs-demo.js
 * AUTH: Brent Ely (https://github.com/gitbrent/)
 * DATE: Feb 27, 2017
 * DESC: Demonstrate PptxGenJS on Node.js
 * REQS: npm 4.x + `npm install pptxgenjs`
 * EXEC: `node nodejs-demo.js`
 */

// ============================================================================
var GIF_ANIM_FIRE = "";
var AUDIO_MP3 = "";
var VIDEO_MP4 = "";

function getTimestamp() {
	var dateNow = new Date();
	var dateMM = dateNow.getMonth() + 1; dateDD = dateNow.getDate(); dateYY = dateNow.getFullYear(), h = dateNow.getHours(); m = dateNow.getMinutes();
	return dateNow.getFullYear() +''+ (dateMM<=9 ? '0' + dateMM : dateMM) +''+ (dateDD<=9 ? '0' + dateDD : dateDD) + (h<=9 ? '0' + h : h) + (m<=9 ? '0' + m : m);
}
// ============================================================================

console.log(`
-------------
STARTING TEST
-------------`);

// STEP 1: Load pptxgenjs and show version to verify everything loaded correctly
//var pptx = require('../dist/pptxgen.js'); // for LOCAL TESTING
var pptx = require("pptxgenjs");
var runEveryTest = require("../examples/pptxgenjs-demo.js");
console.log(` * pptxgenjs version: ${pptx.getVersion()}`); // Loaded okay?

// ============================================================================

function saveCallback(filename) {
	console.log('Good News Everyone!  File created: '+ filename);
}

// ============================================================================

// STEP 2: Run all test funcs
runEveryTest();

// STEP 3: Export giant demo file

// A: Inline save
//pptx.save( 'Node_Demo_NoCallback'+getTimestamp() );

// B: or Save using callback function
//pptx.save( 'Node_Demo_'+getTimestamp(), function(filename){ console.log('Created: '+filename); } );

// C: or use a predefined callback function
pptx.save( 'Node_Demo_Callback_'+getTimestamp(), saveCallback );


// **NOTE** If you continue to use the `pptx` variable, new Slides will be added to the existing set
// Create a new variable or reset `pptx` for an empty Presenation
// EX: pptx = require("pptxgenjs");

// ============================================================================

console.log(`
--------------
TEST COMPLETE!
--------------
 * Files saved to...: ${__dirname}
`);
