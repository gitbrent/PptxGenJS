import Slide from '../slide.js';
const EMU = 914400, CRLF = '\r\n';

function componentToHex(c) {
    var hex = c.toString(16);
    return hex.length == 1 ? "0" + hex : hex;
}

/**
 * Used by {addSlidesForTable} to convert RGB colors from jQuery selectors to Hex for Presentation colors
 */
export function rgbToHex(r, g, b) {
    if (!Number.isInteger(r)) {
        try {
            console.warn('Integer expected!');
        } catch (ex) {}
    }
    return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase();
}

export function inch2Emu(inches) {
    // FIRST: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
    // Any value over 100 damn sure isnt inches, must be EMU already, so just return it
    if (inches > 100) return inches;
    if (typeof inches == 'string') inches = Number(inches.replace(/in*/gi, ''));
    return Math.round(EMU * inches);
}

export function getSizeFromImage(inImgUrl) {
    // A: Create
    var image = new Image();

    // B: Set onload event
    image.onload = function() {
        // FIRST: Check for any errors: This is the best method (try/catch wont work, etc.)
        if (this.width + this.height == 0) {
            return {
                width: 0,
                height: 0
            };
        }
        var obj = {
            width: this.width,
            height: this.height
        };
        return obj;
    };
    image.onerror = function() {
        try {
            console.error('[Error] Unable to load image: ' + inImgUrl);
        } catch (ex) {}
    };

    // C: Load image
    image.src = inImgUrl;
}

export function calcEmuCellHeightForStr(cell, inIntWidthInches) {
    // FORMULA for char-per-inch: (desired chars per line) / (font size [chars-per-inch]) = (reqd print area in inches)
    var GRATIO = 2.61803398875; // "Golden Ratio"
    var intCharPerInch = -1,
        intCalcGratio = 0;

    // STEP 1: Calc chars-per-inch [pitch]
    // SEE: CPL Formula from http://www.pearsonified.com/2012/01/characters-per-line.php
    intCharPerInch = (120 / cell.opts.font_size);

    // STEP 2: Calc line count
    var intLineCnt = Math.floor(cell.text.length / (intCharPerInch * inIntWidthInches));
    if (intLineCnt < 1) intLineCnt = 1; // Dont allow line count to be 0!

    // STEP 3: Calc cell height
    var intCellH = (intLineCnt * ((cell.opts.font_size * 2) / 100));
    if (intLineCnt > 8) intCellH = (intCellH * 0.9);

    // STEP 4: Add cell padding to height
    if (cell.opts.marginPt && Array.isArray(cell.opts.marginPt)) {
        intCellH += (cell.opts.marginPt[0] / ONEPT * (1 / 72)) + (cell.opts.marginPt[2] / ONEPT * (1 / 72));
    } else if (cell.opts.marginPt && Number.isInteger(cell.opts.marginPt)) {
        intCellH += (cell.opts.marginPt / ONEPT * (1 / 72)) + (cell.opts.marginPt / ONEPT * (1 / 72));
    }

    // LAST: Return size
    return inch2Emu(intCellH);
}

export function parseTextToLines(inStr, inFontSize, inWidth) {
    var U = 2.2; // Character Constant thingy
    var CPL = (inWidth / (inFontSize / U));
    var arrLines = [];
    var strCurrLine = '';

    // A: Remove leading/trailing space
    inStr = $.trim(inStr);

    // B: Build line array
    $.each(inStr.split('\n'), function(i, line) {
        $.each(line.split(' '), function(i, word) {
            if (strCurrLine.length + word.length + 1 < CPL) {
                strCurrLine += (word + " ");
            } else {
                if (strCurrLine) arrLines.push(strCurrLine);
                strCurrLine = (word + " ");
            }
        });
        // All words for this line have been exhausted, flush buffer to new line, clear line var
        if (strCurrLine) arrLines.push($.trim(strCurrLine) + CRLF);
        strCurrLine = "";
    });

    // C: Remove trailing linebreak
    arrLines[(arrLines.length - 1)] = $.trim(arrLines[(arrLines.length - 1)]);

    // D: Return lines
    return arrLines;
}

export function getShapeInfo(shapeName) {
    if (!shapeName) return gObjPptxShapes.RECTANGLE;

    if (typeof shapeName == 'object' && shapeName.name && shapeName.displayName && shapeName.avLst) return shapeName;

    if (gObjPptxShapes[shapeName]) return gObjPptxShapes[shapeName];

    var objShape = gObjPptxShapes.filter(function(obj) {
        return obj.name == shapeName || obj.displayName;
    })[0];
    if (typeof objShape !== 'undefined' && objShape != null) return objShape;

    return gObjPptxShapes.RECTANGLE;
}

export function getSmartParseNumber(inVal, inDir, gObjPptx) {
    // FIRST: Convert string numeric value if reqd
    if (typeof inVal == 'string' && !isNaN(Number(inVal))) inVal = Number(inVal);

    // CASE 1: Number in inches
    // Figure any number less than 100 is inches
    if (typeof inVal == 'number' && inVal < 100) return inch2Emu(inVal);

    // CASE 2: Number is already converted to something other than inches
    // Figure any number greater than 100 is not inches! :)  Just return it (its EMU already i guess??)
    if (typeof inVal == 'number' && inVal >= 100) return inVal;

    // CASE 3: Percentage (ex: '50%')
    if (typeof inVal == 'string' && inVal.indexOf('%') > -1) {
        if (inDir && inDir == 'X') return Math.round((parseInt(inVal, 10) / 100) * gObjPptx.pptLayout.width);
        if (inDir && inDir == 'Y') return Math.round((parseInt(inVal, 10) / 100) * gObjPptx.pptLayout.height);
        // Default: Assume width (x/cx)
        return Math.round((parseInt(inVal, 10) / 100) * gObjPptx.pptLayout.width);
    }

    // LAST: Default value
    return 0;
}

export function decodeXmlEntities(inStr) {
    // NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
    if (typeof inStr === 'undefined' || inStr == null) return "";
    return inStr.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/\'/g, '&apos;');
}

export function genXmlColorSelection(color_info, back_info) {
    var outText = '';
    var colorVal;
    var fillType = 'solid';
    var internalElements = '';

    if (back_info) {
        outText += '<p:bg><p:bgPr>';
        outText += genXmlColorSelection(back_info, false);
        outText += '<a:effectLst/>';
        outText += '</p:bgPr></p:bg>';
    }

    if (color_info) {
        if (typeof color_info == 'string') colorVal = color_info;
        else {
            if (color_info.type) fillType = color_info.type;
            if (color_info.color) colorVal = color_info.color;
            if (color_info.alpha) internalElements += '<a:alpha val="' + (100 - color_info.alpha) + '000"/>';
        }

        switch (fillType) {
            case 'solid':
                outText += '<a:solidFill><a:srgbClr val="' + colorVal + '">' + internalElements + '</a:srgbClr></a:solidFill>';
                break;
        }
    }

    return outText;
}

export function convertImgToDataURLviaCanvas(slideRel) {
    // A: Create
    let image = new Image();
    // B: Set onload event
    image.onload = function () {
        // First: Check for any errors: This is the best method (try/catch wont work, etc.)
        if (this.width + this.height == 0) {
            this.onerror();
            return;
        }
        let canvas = document.createElement('CANVAS');
        let ctx = canvas.getContext('2d');
        canvas.height = this.height;
        canvas.width = this.width;
        ctx.drawImage(this, 0, 0);
        // Users running on local machine will get the following error:
        // "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
        // when the canvas.toDataURL call executes below.
        try {
            callbackImgToDataURLDone(canvas.toDataURL(slideRel.type), slideRel);
        } catch (ex) {
            this.onerror();
            console.log("NOTE: Browsers wont let you load/convert local images! (search for --allow-file-access-from-files)");
            return;
        }
        canvas = null;
    };
    image.onerror = () => {
        try {
            console.error('[Error] Unable to load image: ' + slideRel.path);
        } catch (ex) {}
        // Return a predefined "Broken image" graphic so the user will see something on the slide
        callbackImgToDataURLDone('data:image/png;base64,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', slideRel);
    };
    // C: Load image
    image.src = slideRel.path;
}

function callbackImgToDataURLDone(inStr, slideRel) {
    var intEmpty = 0;

    // STEP 1: Store base64 data for this image
    slideRel.data = inStr;

    // STEP 2: Call export function once all async processes have completed
    $.each(Slide.gObjPptx.slides, (i, slide) => {
        $.each(slide.rels, (i, rel) => {
            if (rel.path == slideRel.path) rel.data = inStr;
            if (!rel.data) intEmpty++;
        });
    });

    // STEP 3: Continue export process
    if (intEmpty == 0) this; //.save();
}