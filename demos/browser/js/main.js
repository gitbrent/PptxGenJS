import {
	doAppStart, execGenSlidesFunc, runAllDemos,
	table2slides1, table2slides2, table2slidesDemoForTab,
	doRunBasicDemo, doRunSandboxDemo, buildDataTable, padDataTable
} from './browser.js';

// STEP 1: Add event listeners to "run demo" buttons
document.getElementById('btnRunAllDemos').addEventListener('click', () => runAllDemos());
document.getElementById('btnRunBasicDemo').addEventListener('click', () => doRunBasicDemo());
document.getElementById('btnRunSandboxDemo').addEventListener('click', () => doRunSandboxDemo());
document.getElementById('btnGenFunc_Chart').addEventListener('click', () => execGenSlidesFunc('Chart'));
document.getElementById('btnGenFunc_Image').addEventListener('click', () => execGenSlidesFunc('Image'));
document.getElementById('btnGenFunc_Media').addEventListener('click', () => execGenSlidesFunc('Media'));
document.getElementById('btnGenFunc_Shape').addEventListener('click', () => execGenSlidesFunc('Shape'));
document.getElementById('btnGenFunc_Text').addEventListener('click', () => execGenSlidesFunc('Text'));
document.getElementById('btnGenFunc_Table').addEventListener('click', () => execGenSlidesFunc('Table'));
document.getElementById('btnGenFunc_Master').addEventListener('click', () => execGenSlidesFunc('Master'));

// STEP 2: HTML-to-PPTX: Dynamic Table input handlers
document.getElementById('table2slides1').addEventListener('click', () => table2slides1());
document.getElementById('table2slides2').addEventListener('click', () => table2slides2(false));
document.getElementById('table2slides3').addEventListener('click', () => table2slides2(true));
document.getElementById('tab2slides_tabNoStyle').addEventListener('click', () => table2slidesDemoForTab('tabNoStyle'));
document.getElementById('tab2slides_tabInheritStyle').addEventListener('click', () => table2slidesDemoForTab('tabInheritStyle'));
document.getElementById('tab2slides_tabColspan').addEventListener('click', () => table2slidesDemoForTab('tabColspan'));
document.getElementById('tab2slides_tabRowspan').addEventListener('click', () => table2slidesDemoForTab('tabRowspan'));
document.getElementById('tab2slides_tabRowColspan').addEventListener('click', () => table2slidesDemoForTab('tabRowColspan'));
document.getElementById('tab2slides_tabLotsOfLines').addEventListener('click', () => table2slidesDemoForTab('tabLotsOfLines', { verbose: false }));
document.getElementById('tab2slides_tabLargeCellText').addEventListener('click', () => table2slidesDemoForTab('tabLargeCellText', { verbose: false }));
document.getElementById('numTab2SlideRows').addEventListener('change', () => buildDataTable());
document.getElementById('numTab2Padding').addEventListener('change', () => padDataTable());

// LAST: START!
document.addEventListener('DOMContentLoaded', () => doAppStart());
