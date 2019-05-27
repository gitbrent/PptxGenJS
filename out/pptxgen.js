/*\
|*|  :: pptxgen.js ::
|*|
|*|  JavaScript framework that creates PowerPoint (pptx) presentations
|*|  https://github.com/gitbrent/PptxGenJS
|*|
|*|  This framework is released under the MIT Public License (MIT)
|*|
|*|  PptxGenJS (C) 2015-2019 Brent Ely -- https://github.com/gitbrent
|*|
|*|  Some code derived from the OfficeGen project:
|*|  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
|*|
|*|  Permission is hereby granted, free of charge, to any person obtaining a copy
|*|  of this software and associated documentation files (the "Software"), to deal
|*|  in the Software without restriction, including without limitation the rights
|*|  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
|*|  copies of the Software, and to permit persons to whom the Software is
|*|  furnished to do so, subject to the following conditions:
|*|
|*|  The above copyright notice and this permission notice shall be included in all
|*|  copies or substantial portions of the Software.
|*|
|*|  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
|*|  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
|*|  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
|*|  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
|*|  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
|*|  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
|*|  SOFTWARE.
\*/
/*
    PPTX Units are "DXA" (except for font sizing)
    ....: There are 1440 DXA per inch. 1 inch is 72 points. 1 DXA is 1/20th's of a point (20 DXA is 1 point).
    ....: There is also something called EMU's (914400 EMUs is 1 inch, 12700 EMUs is 1pt).
    SEE: https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
    |
    OBJECT LAYOUTS: 16x9 (10" x 5.625"), 16x10 (10" x 6.25"), 4x3 (10" x 7.5"), Wide (13.33" x 7.5") and Custom (any size)
    |
    REFS:
    * "Structure of a PresentationML document (Open XML SDK)"
    * @see: https://msdn.microsoft.com/en-us/library/office/gg278335.aspx
    * TableStyleId enumeration
    * @see: https://msdn.microsoft.com/en-us/library/office/hh273476(v=office.14).aspx
*/
import { EMU, ONEPT, CRLF, DEF_CELL_MARGIN_PT, DEF_FONT_COLOR, DEF_FONT_SIZE, DEF_SLIDE_MARGIN_IN, MASTER_OBJECTS, IMG_BROKEN, IMG_PLAYBTN, SLIDE_OBJECT_TYPES, } from './enums';
import { getSmartParseNumber, inch2Emu, rgbToHex } from './utils';
import * as genXml from './gen-xml';
var PptxGenJS = /** @class */ (function () {
    function PptxGenJS() {
        var _this = this;
        this._version = '3.0.0';
        this.NODEJS = false;
        /**
         * DESC: Export the .pptx file
         */
        this.doExportPresentation = function (outputType) {
            var arrChartPromises = [];
            var intSlideNum = 0;
            // STEP 1: Create new JSZip file
            var zip = new _this.JSZip();
            // STEP 2: Add all required folders and files
            zip.folder('_rels');
            zip.folder('docProps');
            zip.folder('ppt').folder('_rels');
            zip.folder('ppt/charts').folder('_rels');
            zip.folder('ppt/embeddings');
            zip.folder('ppt/media');
            zip.folder('ppt/slideLayouts').folder('_rels');
            zip.folder('ppt/slideMasters').folder('_rels');
            zip.folder('ppt/slides').folder('_rels');
            zip.folder('ppt/theme');
            zip.folder('ppt/notesMasters').folder('_rels');
            zip.folder('ppt/notesSlides').folder('_rels');
            //
            zip.file('[Content_Types].xml', genXml.makeXmlContTypes(_this.slides, _this.slideLayouts, _this.masterSlide));
            zip.file('_rels/.rels', genXml.makeXmlRootRels());
            zip.file('docProps/app.xml', genXml.makeXmlApp(_this.slides, _this.company));
            zip.file('docProps/core.xml', genXml.makeXmlCore(_this.title, _this.subject, _this.author, _this.revision));
            zip.file('ppt/_rels/presentation.xml.rels', genXml.makeXmlPresentationRels(_this.slides));
            //
            zip.file('ppt/theme/theme1.xml', genXml.makeXmlTheme());
            zip.file('ppt/presentation.xml', genXml.makeXmlPresentation(_this.slides, _this.pptLayout));
            zip.file('ppt/presProps.xml', genXml.makeXmlPresProps());
            zip.file('ppt/tableStyles.xml', genXml.makeXmlTableStyles());
            zip.file('ppt/viewProps.xml', genXml.makeXmlViewProps());
            // Create a Layout/Master/Rel/Slide file for each SLIDE
            for (var idx = 1; idx <= _this.slideLayouts.length; idx++) {
                zip.file('ppt/slideLayouts/slideLayout' + idx + '.xml', genXml.makeXmlLayout(_this.slideLayouts[idx - 1]));
                zip.file('ppt/slideLayouts/_rels/slideLayout' + idx + '.xml.rels', genXml.makeXmlSlideLayoutRel(idx, _this.slideLayouts));
            }
            for (var idx = 0; idx < _this.slides.length; idx++) {
                intSlideNum++;
                zip.file('ppt/slides/slide' + intSlideNum + '.xml', genXml.makeXmlSlide(_this.slides[idx]));
                zip.file('ppt/slides/_rels/slide' + intSlideNum + '.xml.rels', genXml.makeXmlSlideRel(_this.slides, _this.slideLayouts, intSlideNum));
                // Here we will create all slide notes related items. Notes of empty strings
                // are created for slides which do not have notes specified, to keep track of _rels.
                zip.file('ppt/notesSlides/notesSlide' + intSlideNum + '.xml', genXml.makeXmlNotesSlide(_this.slides[idx]));
                zip.file('ppt/notesSlides/_rels/notesSlide' + intSlideNum + '.xml.rels', genXml.makeXmlNotesSlideRel(intSlideNum));
            }
            zip.file('ppt/slideMasters/slideMaster1.xml', genXml.makeXmlMaster(_this.masterSlide, _this.slideLayouts));
            zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', genXml.makeXmlMasterRel(_this.masterSlide, _this.slideLayouts));
            zip.file('ppt/notesMasters/notesMaster1.xml', genXml.makeXmlNotesMaster());
            zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', genXml.makeXmlNotesMasterRel());
            // Create all Rels (images, media, chart data)
            _this.slideLayouts.forEach(function (layout) {
                this.createMediaFiles(layout, zip, arrChartPromises);
            });
            _this.slides.forEach(function (slide) {
                this.createMediaFiles(slide, zip, arrChartPromises);
            });
            _this.createMediaFiles(_this.masterSlide, zip, arrChartPromises);
            // STEP 3: Wait for Promises (if any) then generate the PPTX file
            Promise.all(arrChartPromises)
                .then(function () {
                var strExportName = _this.fileName.toLowerCase().indexOf('.ppt') > -1 ? _this.fileName : _this.fileName + _this.fileExtn;
                if (outputType) {
                    zip.generateAsync({ type: outputType }).then(function () { return _this.saveCallback; });
                }
                else if (_this.NODEJS && !_this.isBrowser) {
                    if (_this.saveCallback) {
                        if (strExportName.indexOf('http') == 0) {
                            zip.generateAsync({ type: 'nodebuffer' }).then(function (content) {
                                this.saveCallback(content);
                            });
                        }
                        else {
                            zip.generateAsync({ type: 'nodebuffer' }).then(function (content) {
                                this.fs.writeFile(strExportName, content, function () {
                                    this.saveCallback(strExportName);
                                });
                            });
                        }
                    }
                    else {
                        // Starting in late 2017 (Node ~8.9.1), `fs` requires a callback so use a dummy func
                        zip.generateAsync({ type: 'nodebuffer' }).then(function (content) {
                            this.fs.writeFile(strExportName, content, function () { });
                        });
                    }
                }
                else {
                    zip.generateAsync({ type: 'blob' }).then(function (content) {
                        this.writeFileToBrowser(strExportName, content);
                    });
                }
            })
                .catch(function (strErr) {
                console.error(strErr);
            });
        };
        this.writeFileToBrowser = function (strExportName, content) {
            // STEP 1: Create element
            var a = document.createElement('a');
            a.setAttribute('style', 'display:none;');
            document.body.appendChild(a);
            // STEP 2: Download file to browser
            // DESIGN: Use `createObjectURL()` (or MS-specific func for IE11) to D/L files in client browsers (FYI: synchronously executed)
            if (window.navigator.msSaveOrOpenBlob) {
                // REF: https://docs.microsoft.com/en-us/microsoft-edge/dev-guide/html5/file-api/blob
                var blobObject_1 = new Blob([content]);
                jQuery(a).click(function () {
                    window.navigator.msSaveOrOpenBlob(blobObject_1, strExportName);
                });
                a.click();
                // Clean-up
                document.body.removeChild(a);
                // LAST: Callback (if any)
                if (_this.saveCallback)
                    _this.saveCallback(strExportName);
            }
            else if (window.URL.createObjectURL) {
                var blob = new Blob([content], { type: 'octet/stream' });
                var url = window.URL.createObjectURL(blob);
                a.href = url;
                a.download = strExportName;
                a.click();
                // Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
                setTimeout(function () {
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                }, 100);
                // LAST: Callback (if any)
                if (_this.saveCallback)
                    _this.saveCallback(strExportName);
            }
            // STEP 3: Clear callback func post-save
            _this.saveCallback = null;
        };
        this.createMediaFiles = function (layout, zip, chartPromises) {
            layout.relsChart.forEach(function (rel) { return chartPromises.push(genXml.gObjPptxGenerators.createExcelWorksheet(rel, zip)); });
            layout.relsMedia.forEach(function (rel) {
                if (rel.type != 'online' && rel.type != 'hyperlink') {
                    // A: Loop vars
                    var data = rel.data;
                    // B: Users will undoubtedly pass various string formats, so correct prefixes as needed
                    if (data.indexOf(',') == -1 && data.indexOf(';') == -1)
                        data = 'image/png;base64,' + data;
                    else if (data.indexOf(',') == -1)
                        data = 'image/png;base64,' + data;
                    else if (data.indexOf(';') == -1)
                        data = 'image/png;' + data;
                    // C: Add media
                    zip.file(rel.Target.replace('..', 'ppt'), data.split(',').pop(), { base64: true });
                }
            });
        };
        this.addPlaceholdersToSlides = function (slide) {
            // Add all placeholders on this Slide that dont already exist
            slide.layoutObj.data.forEach(function (slideLayoutObj) {
                if (slideLayoutObj.type === MASTER_OBJECTS.placeholder) {
                    // A: Search for this placeholder on Slide before we add
                    // NOTE: Check to ensure a placeholder does not already exist on the Slide
                    // They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
                    if (slide.data.filter(function (slideObj) {
                        return slideObj.options && slideObj.options.placeholder == slideLayoutObj.options.placeholderName;
                    }).length == 0) {
                        genXml.gObjPptxGenerators.addTextDefinition('', { placeholder: slideLayoutObj.options.placeholderName }, slide, false);
                    }
                }
            });
        };
        // IMAGE METHODS:
        this.getSizeFromImage = function (inImgUrl) {
            if (_this.NODEJS) {
                try {
                    var dimensions = _this.sizeOf(inImgUrl);
                    return { width: dimensions.width, height: dimensions.height };
                }
                catch (ex) {
                    console.error('ERROR: Unable to read image: ' + inImgUrl);
                    return { width: 0, height: 0 };
                }
            }
            // A: Create
            var image = new Image();
            // B: Set onload event
            image.onload = function () {
                // FIRST: Check for any errors: This is the best method (try/catch wont work, etc.)
                if (image.width + image.height == 0) {
                    return { width: 0, height: 0 };
                }
                var obj = { width: image.width, height: image.height };
                return obj;
            };
            image.onerror = function () {
                try {
                    console.error('[Error] Unable to load image: ' + inImgUrl);
                }
                catch (ex) { }
            };
            // C: Load image
            image.src = inImgUrl;
        };
        /* Encode Image/Audio/Video into base64 */
        this.encodeSlideMediaRels = function (layout, arrRelsDone) {
            var intRels = 0;
            layout.rels.forEach(function (rel) {
                // Read and Encode each media lacking `data` into base64 (for use in export)
                if (rel.type != 'online' && rel.type != 'chart' && !rel.data && arrRelsDone.indexOf(rel.path) == -1) {
                    // Node local-file encoding is syncronous, so we can load all images here, then call export with a callback (if any)
                    if (this.NODEJS && rel.path.indexOf('http') != 0) {
                        try {
                            var bitmap = this.fs.readFileSync(rel.path);
                            rel.data = Buffer.from(bitmap).toString('base64');
                        }
                        catch (ex) {
                            console.error('ERROR....: Unable to read media: "' + rel.path + '"');
                            console.error('DETAILS..: ' + ex);
                            rel.data = IMG_BROKEN;
                        }
                    }
                    else if (this.NODEJS && rel.path.indexOf('http') == 0) {
                        intRels++;
                        this.convertRemoteMediaToDataURL(rel);
                        arrRelsDone.push(rel.path);
                    }
                    else {
                        intRels++;
                        this.convertImgToDataURL(rel);
                        arrRelsDone.push(rel.path);
                    }
                }
                else if (rel.isSvgPng && rel.data && rel.data.toLowerCase().indexOf('image/svg') > -1) {
                    // The SVG base64 must be converted to PNG SVG before export
                    intRels++;
                    this.callbackImgToDataURLDone(rel.data, rel);
                    arrRelsDone.push(rel.path);
                }
            });
            return intRels;
        };
        /* `FileReader()` + `readAsDataURL` = Ablity to read any file into base64! */
        this.convertImgToDataURL = function (slideRel) {
            var xhr = new XMLHttpRequest();
            xhr.onload = function () {
                var reader = new FileReader();
                reader.onloadend = function () {
                    _this.callbackImgToDataURLDone(reader.result, slideRel);
                };
                reader.readAsDataURL(xhr.response);
            };
            xhr.onerror = function (ex) {
                // TODO: xhr.error/catch whatever! then return
                console.error('Unable to load image: "' + slideRel.path);
                console.error(ex || '');
                // Return a predefined "Broken image" graphic so the user will see something on the slide
                _this.callbackImgToDataURLDone(IMG_BROKEN, slideRel);
            };
            xhr.open('GET', slideRel.path);
            xhr.responseType = 'blob';
            xhr.send();
        };
        /* Node equivalent of `convertImgToDataURL()`: Use https to fetch, then use Buffer to encode to base64 */
        this.convertRemoteMediaToDataURL = function (slideRel) {
            _this.https.get(slideRel.path, function (res) {
                var rawData = '';
                res.setEncoding('binary'); // IMPORTANT: Only binary encoding works
                res.on('data', function (chunk) {
                    rawData += chunk;
                });
                res.on('end', function () {
                    var data = Buffer.from(rawData, 'binary').toString('base64');
                    _this.callbackImgToDataURLDone(data, slideRel);
                });
                res.on('error', function (e) {
                    // TODO-3:
                    ///reject(e);
                });
            });
        };
        /* browser: Convert SVG-base64 data to PNG-base64 */
        this.convertSvgToPngViaCanvas = function (slideRel) {
            // A: Create
            var image = new Image();
            // B: Set onload event
            image.onload = function () {
                // First: Check for any errors: This is the best method (try/catch wont work, etc.)
                if (image.width + image.height == 0) {
                    image.onerror('h/w=0');
                    return;
                }
                var canvas = document.createElement('CANVAS');
                var ctx = canvas.getContext('2d');
                canvas.width = image.width;
                canvas.height = image.height;
                ctx.drawImage(image, 0, 0);
                // Users running on local machine will get the following error:
                // "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
                // when the canvas.toDataURL call executes below.
                try {
                    _this.callbackImgToDataURLDone(canvas.toDataURL(slideRel.type), slideRel);
                }
                catch (ex) {
                    image.onerror(ex);
                    return;
                }
                canvas = null;
            };
            image.onerror = function (ex) {
                console.error(ex || '');
                // Return a predefined "Broken image" graphic so the user will see something on the slide
                this.callbackImgToDataURLDone(IMG_BROKEN, slideRel);
            };
            // C: Load image
            image.src = slideRel.data; // use pre-encoded SVG base64 data
        };
        this.callbackImgToDataURLDone = function (base64Data, slideRel) {
            // SVG images were retrieved via `convertImgToDataURL()`, but have to be encoded to PNG now
            if (slideRel.isSvgPng && typeof base64Data === 'string' && base64Data.indexOf('image/svg') > -1) {
                // Pass the SVG XML as base64 for conversion to PNG
                slideRel.data = base64Data;
                if (_this.NODEJS)
                    console.log('SVG is not supported in Node');
                else
                    _this.convertSvgToPngViaCanvas(slideRel);
                return;
            }
            var intEmpty = 0;
            var funcCallback = function (rel) {
                if (rel.path == slideRel.path)
                    rel.data = base64Data;
                if (!rel.data)
                    intEmpty++;
            };
            // STEP 1: Set data for this rel, count outstanding
            _this.slides.forEach(function (slide) {
                slide.rels.forEach(funcCallback);
            });
            _this.slideLayouts.forEach(function (layout) {
                layout.rels.forEach(funcCallback);
            });
            _this.masterSlide.rels.forEach(funcCallback);
            // STEP 2: Continue export process if all rels have base64 `data` now
            if (intEmpty == 0)
                _this.doExportPresentation();
        };
        /**
         * Magic happens here
         */
        this.parseTextToLines = function (cell, inWidth) {
            var CHAR = 2.2 + (cell.opts && cell.opts.lineWeight ? cell.opts.lineWeight : 0); // Character Constant (An approximation of the Golden Ratio)
            var CPL = (inWidth * EMU) / ((cell.opts.fontSize || DEF_FONT_SIZE) / CHAR); // Chars-Per-Line
            var arrLines = [];
            var strCurrLine = '';
            // Allow a single space/whitespace as cell text
            if (cell.text && cell.text.trim() == '')
                return [' '];
            // A: Remove leading/trailing space
            var inStr = (cell.text || '').toString().trim();
            // B: Build line array
            jQuery.each(inStr.split('\n'), function (i, line) {
                jQuery.each(line.split(' '), function (i, word) {
                    if (strCurrLine.length + word.length + 1 < CPL) {
                        strCurrLine += word + ' ';
                    }
                    else {
                        if (strCurrLine)
                            arrLines.push(strCurrLine);
                        strCurrLine = word + ' ';
                    }
                });
                // All words for this line have been exhausted, flush buffer to new line, clear line var
                if (strCurrLine)
                    arrLines.push(jQuery.trim(strCurrLine) + CRLF);
                strCurrLine = '';
            });
            // C: Remove trailing linebreak
            arrLines[arrLines.length - 1] = jQuery.trim(arrLines[arrLines.length - 1]);
            // D: Return lines
            return arrLines;
        };
        /**
         * Magic happens here
         */
        this.getSlidesForTableRows = function (inArrRows, opts) {
            var LINEH_MODIFIER = 1.9;
            var opts = opts || {};
            var arrInchMargins = DEF_SLIDE_MARGIN_IN; // (0.5" on all sides)
            var arrObjTabHeadRows = opts.arrObjTabHeadRows || [], arrObjTabBodyRows = [], arrObjTabFootRows = [];
            var arrObjSlides = [], arrRows = [], currRow = [], numCols = 0;
            var intTabW = 0, emuTabCurrH = 0, emuSlideTabW = EMU * 1, emuSlideTabH = EMU * 1;
            if (opts.debug)
                console.log('------------------------------------');
            if (opts.debug)
                console.log('opts.w ............. = ' + (opts.w || '').toString());
            if (opts.debug)
                console.log('opts.colW .......... = ' + (opts.colW || '').toString());
            if (opts.debug)
                console.log('opts.slideMargin ... = ' + (opts.slideMargin || '').toString());
            // NOTE: Use default size as zero cell margin is causing our tables to be too large and touch bottom of slide!
            if (!opts.slideMargin && opts.slideMargin != 0)
                opts.slideMargin = DEF_SLIDE_MARGIN_IN[0];
            // STEP 1: Calc margins/usable space
            if (opts.slideMargin || opts.slideMargin == 0) {
                if (Array.isArray(opts.slideMargin))
                    arrInchMargins = opts.slideMargin;
                else if (!isNaN(opts.slideMargin))
                    arrInchMargins = [opts.slideMargin, opts.slideMargin, opts.slideMargin, opts.slideMargin];
            }
            else if (opts && opts.master && opts.master.margin) {
                if (Array.isArray(opts.master.margin))
                    arrInchMargins = opts.master.margin;
                else if (!isNaN(opts.master.margin))
                    arrInchMargins = [opts.master.margin, opts.master.margin, opts.master.margin, opts.master.margin];
            }
            // STEP 2: Calc number of columns
            // NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
            // ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
            inArrRows[0].forEach(function (cell, idx) {
                if (!cell)
                    cell = {};
                var cellOpts = cell.options || cell.opts || null;
                numCols += cellOpts && cellOpts.colspan ? cellOpts.colspan : 1;
            });
            if (opts.debug)
                console.log('arrInchMargins ..... = ' + arrInchMargins.toString());
            if (opts.debug)
                console.log('numCols ............ = ' + numCols);
            // Calc opts.w if we can
            if (!opts.w && opts.colW) {
                if (Array.isArray(opts.colW))
                    opts.colW.forEach(function (val, idx) {
                        opts.w += val;
                    });
                else {
                    opts.w = opts.colW * numCols;
                }
            }
            // STEP 2: Calc usable space/table size now that we have usable space calc'd
            emuSlideTabW = opts.w ? inch2Emu(opts.w) : _this.pptLayout.width - inch2Emu((opts.x || arrInchMargins[1]) + arrInchMargins[3]);
            if (opts.debug)
                console.log('emuSlideTabW (in) ........ = ' + (emuSlideTabW / EMU).toFixed(1));
            if (opts.debug)
                console.log('this.pptLayout.h ..... = ' + _this.pptLayout.height / EMU);
            // STEP 3: Calc column widths if needed so we can subsequently calc lines (we need `emuSlideTabW`!)
            if (!opts.colW || !Array.isArray(opts.colW)) {
                if (opts.colW && !isNaN(Number(opts.colW))) {
                    var arrColW = [];
                    inArrRows[0].forEach(function (cell, idx) {
                        arrColW.push(opts.colW);
                    });
                    opts.colW = [];
                    arrColW.forEach(function (val, idx) {
                        opts.colW.push(val);
                    });
                }
                // No column widths provided? Then distribute cols.
                else {
                    opts.colW = [];
                    for (var iCol = 0; iCol < numCols; iCol++) {
                        opts.colW.push(emuSlideTabW / EMU / numCols);
                    }
                }
            }
            // STEP 4: Iterate over each line and perform magic =========================
            // NOTE: inArrRows will be an array of {text:'', opts{}} whether from `addSlidesForTable()` or `.addTable()`
            inArrRows.forEach(function (row, iRow) {
                // A: Reset ROW variables
                var arrCellsLines = [], arrCellsLineHeights = [], emuRowH = 0, intMaxLineCnt = 0, intMaxColIdx = 0;
                // B: Calc usable vertical space/table height
                // NOTE: Use margins after the first Slide (dont re-use opt.y - it could've been halfway down the page!) (ISSUE#43,ISSUE#47,ISSUE#48)
                if (arrObjSlides.length > 0) {
                    emuSlideTabH = this.pptLayout.height - inch2Emu((opts.y / EMU < arrInchMargins[0] ? opts.y / EMU : arrInchMargins[0]) + arrInchMargins[2]);
                    // Use whichever is greater: area between margins or the table H provided (dont shrink usable area - the whole point of over-riding X on paging is to *increarse* usable space)
                    if (emuSlideTabH < opts.h)
                        emuSlideTabH = opts.h;
                }
                else
                    emuSlideTabH = opts.h ? opts.h : this.pptLayout.height - inch2Emu((opts.y / EMU || arrInchMargins[0]) + arrInchMargins[2]);
                if (opts.debug)
                    console.log('* Slide ' + arrObjSlides.length + ': emuSlideTabH (in) ........ = ' + (emuSlideTabH / EMU).toFixed(1));
                // C: Parse and store each cell's text into line array (**MAGIC HAPPENS HERE**)
                row.forEach(function (cell, iCell) {
                    // FIRST: REALITY-CHECK:
                    if (!cell)
                        cell = {};
                    // DESIGN: Cells are henceforth {objects} with `text` and `opts`
                    var lines = [];
                    // 1: Cleanse data
                    if (!isNaN(cell) || typeof cell === 'string') {
                        // Grab table formatting `opts` to use here so text style/format inherits as it should
                        cell = { text: cell.toString(), opts: opts };
                    }
                    else if (typeof cell === 'object') {
                        // ARG0: `text`
                        if (typeof cell.text === 'number')
                            cell.text = cell.text.toString();
                        else if (typeof cell.text === 'undefined' || cell.text == null)
                            cell.text = '';
                        // ARG1: `options`
                        var opt = cell.options || cell.opts || {};
                        cell.opts = opt;
                    }
                    // Capture some table options for use in other functions
                    cell.opts.lineWeight = opts.lineWeight;
                    // 2: Create a cell object for each table column
                    currRow.push({ text: '', opts: cell.opts });
                    // 3: Parse cell contents into lines (**MAGIC HAPPENSS HERE**)
                    var lines = this.parseTextToLines(cell, opts.colW[iCell] / ONEPT);
                    arrCellsLines.push(lines);
                    //if (opts.debug) console.log('Cell:'+iCell+' - lines:'+lines.length);
                    // 4: Keep track of max line count within all row cells
                    if (lines.length > intMaxLineCnt) {
                        intMaxLineCnt = lines.length;
                        intMaxColIdx = iCell;
                    }
                    var lineHeight = inch2Emu(((cell.opts.fontSize || opts.fontSize || DEF_FONT_SIZE) * LINEH_MODIFIER) / 100);
                    // NOTE: Exempt cells with `rowspan` from increasing lineHeight (or we could create a new slide when unecessary!)
                    if (cell.opts && cell.opts.rowspan)
                        lineHeight = 0;
                    // 5: Add cell margins to lineHeight (if any)
                    if (cell.opts.margin) {
                        if (cell.opts.margin[0])
                            lineHeight += (cell.opts.margin[0] * ONEPT) / intMaxLineCnt;
                        if (cell.opts.margin[2])
                            lineHeight += (cell.opts.margin[2] * ONEPT) / intMaxLineCnt;
                    }
                    // Add to array
                    arrCellsLineHeights.push(Math.round(lineHeight));
                });
                // D: AUTO-PAGING: Add text one-line-a-time to this row's cells until: lines are exhausted OR table H limit is hit
                for (var idx = 0; idx < intMaxLineCnt; idx++) {
                    // 1: Add the current line to cell
                    for (var col = 0; col < arrCellsLines.length; col++) {
                        // A: Commit this slide to Presenation if table Height limit is hit
                        if (emuTabCurrH + arrCellsLineHeights[intMaxColIdx] > emuSlideTabH) {
                            if (opts.debug)
                                console.log('--------------- New Slide Created ---------------');
                            if (opts.debug)
                                console.log(' (calc) ' +
                                    (emuTabCurrH / EMU).toFixed(1) +
                                    '+' +
                                    (arrCellsLineHeights[intMaxColIdx] / EMU).toFixed(1) +
                                    ' > ' +
                                    (emuSlideTabH / EMU).toFixed(1));
                            if (opts.debug)
                                console.log('--------------- New Slide Created ---------------');
                            // 1: Add the current row to table
                            // NOTE: Edge cases can occur where we create a new slide only to have no more lines
                            // ....: and then a blank row sits at the bottom of a table!
                            // ....: Hence, we verify all cells have text before adding this final row.
                            jQuery.each(currRow, function (i, cell) {
                                if (cell.text.length > 0) {
                                    // IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
                                    arrRows.push(jQuery.extend(true, [], currRow));
                                    return false; // break out of .each loop
                                }
                            });
                            // 2: Add new Slide with current array of table rows
                            arrObjSlides.push(jQuery.extend(true, [], arrRows));
                            // 3: Empty rows for new Slide
                            arrRows.length = 0;
                            // 4: Reset current table height for new Slide
                            emuTabCurrH = 0; // This row's emuRowH w/b added below
                            // 5: Empty current row's text (continue adding lines where we left off below)
                            jQuery.each(currRow, function (i, cell) {
                                cell.text = '';
                            });
                            // 6: Auto-Paging Options: addHeaderToEach
                            if (opts.addHeaderToEach && arrObjTabHeadRows)
                                arrRows = arrRows.concat(arrObjTabHeadRows);
                        }
                        // B: Add next line of text to this cell
                        if (arrCellsLines[col][idx])
                            currRow[col].text += arrCellsLines[col][idx];
                    }
                    // 2: Add this new rows H to overall (use cell with the most lines as the determiner for overall row Height)
                    emuTabCurrH += arrCellsLineHeights[intMaxColIdx];
                }
                if (opts.debug)
                    console.log('-> ' + iRow + ' row done!');
                if (opts.debug)
                    console.log('-> emuTabCurrH (in) . = ' + (emuTabCurrH / EMU).toFixed(1));
                // E: Flush row buffer - Add the current row to table, then truncate row cell array
                // IMPORTANT: use jQuery extend (deep copy) or cell will mutate!!
                if (currRow.length)
                    arrRows.push(jQuery.extend(true, [], currRow));
                currRow.length = 0;
            });
            // STEP 4-2: Flush final row buffer to slide
            arrObjSlides.push(jQuery.extend(true, [], arrRows));
            // LAST:
            if (opts.debug) {
                console.log('arrObjSlides count = ' + arrObjSlides.length);
                console.log(arrObjSlides);
            }
            return arrObjSlides;
        };
        /**
         * Expose a couple private helper functions from above
         */
        this.inch2Emu = function () { return inch2Emu; };
        this.rgbToHex = function () { return rgbToHex; };
        // Determine Environment
        if (typeof module !== 'undefined' && module.exports && typeof require === 'function' && typeof window === 'undefined') {
            try {
                require.resolve('fs');
                this.NODEJS = true;
            }
            catch (ex) {
                this.NODEJS = false;
            }
        }
        this.LAYOUTS = {
            LAYOUT_4x3: { name: 'screen4x3', width: 9144000, height: 6858000 },
            LAYOUT_16x9: { name: 'screen16x9', width: 9144000, height: 5143500 },
            LAYOUT_16x10: { name: 'screen16x10', width: 9144000, height: 5715000 },
            LAYOUT_WIDE: { name: 'custom', width: 12192000, height: 6858000 },
            LAYOUT_USER: { name: 'custom', width: 12192000, height: 6858000 },
        };
        // Core
        this._author = 'PptxGenJS';
        this._company = 'PptxGenJS';
        this._revision = '1';
        this._subject = 'PptxGenJS Presentation';
        this._title = 'PptxGenJS Presentation';
        // PptxGenJS props
        this._pptLayout = this.LAYOUTS['LAYOUT_16x9'];
        this._rtlMode = false;
        this._isBrowser = false;
        this.fileName = 'Presentation';
        this.fileExtn = '.pptx';
        //this.saveCallback = null
        //
        this.masterSlide = {
            slide: null,
            numb: null,
            name: null,
            layoutName: null,
            layoutObj: null,
            data: [],
            rels: [],
            slideNumberObj: null,
        };
        this.slides = [];
        this.slideLayouts = [
            {
                name: 'BLANK',
                slide: null,
                data: [],
                rels: [],
                margin: DEF_SLIDE_MARGIN_IN,
                slideNumberObj: null,
            },
        ];
    }
    Object.defineProperty(PptxGenJS.prototype, "version", {
        get: function () {
            return this._version;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "author", {
        get: function () {
            return this._author;
        },
        set: function (value) {
            this._author = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "company", {
        get: function () {
            return this._company;
        },
        set: function (value) {
            this._company = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "revision", {
        get: function () {
            return this._revision;
        },
        set: function (value) {
            this._revision = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "subject", {
        get: function () {
            return this._subject;
        },
        set: function (value) {
            this._subject = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "title", {
        get: function () {
            return this._title;
        },
        set: function (value) {
            this._title = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "rtlMode", {
        get: function () {
            return this._rtlMode;
        },
        set: function (value) {
            this._rtlMode = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "pptLayout", {
        get: function () {
            return this._pptLayout;
        },
        set: function (value) {
            this._pptLayout = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PptxGenJS.prototype, "isBrowser", {
        get: function () {
            return this._isBrowser;
        },
        set: function (value) {
            this._isBrowser = value;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Save (export) the Presentation .pptx file
     * @param {string} `inStrExportName` - Filename to use for the export
     * @param {Function} `funcCallback` - Callback function to be called when export is complete
     * @param {JSZIP_OUTPUT_TYPE} `outputType` - JSZip output type
     */
    PptxGenJS.prototype.save = function (inStrExportName, funcCallback, outputType) {
        var intRels = 0, arrRelsDone = [];
        // STEP 1: Add empty placeholder objects to slides that don't already have them
        this.slides.forEach(function (slide) {
            if (slide.layoutObj)
                this.addPlaceholdersToSlides(slide);
        });
        // STEP 2: Set export properties
        if (funcCallback)
            this.saveCallback = funcCallback;
        if (inStrExportName)
            this.fileName = inStrExportName;
        // STEP 3: Read/Encode Images
        // PERF: Only send unique paths for encoding (encoding func will find and fill *ALL* matching paths across the Presentation)
        // A: Slide rels
        this.slides.forEach(function (slide, idx) {
            intRels += this.encodeSlideMediaRels(slide, arrRelsDone);
        });
        // B: Layout rels
        this.slideLayouts.forEach(function (layout, idx) {
            intRels += this.encodeSlideMediaRels(layout, arrRelsDone);
        });
        // C: Master Slide rels
        intRels += this.encodeSlideMediaRels(this.masterSlide, arrRelsDone);
        // STEP 4: Export now if there's no images to encode (otherwise, last async imgConvert call above will call exportFile)
        if (intRels == 0)
            this.doExportPresentation(outputType);
    };
    /**
     * Add a new Slide to the Presentation
     * @param {string} inMasterName - Name of Master Slide
     * @returns {ISlide[]} slideObj - The new Slide object
     */
    PptxGenJS.prototype.addNewSlide = function (inMasterName) {
        var _this = this;
        var slideObj = {
            getPageNumber: null,
            slideNumber: null,
            addChart: null,
            addImage: null,
            addMedia: null,
            addNotes: null,
            addShape: null,
            addTable: null,
            addText: null,
        };
        var slideNum = this.slides.length;
        var pageNum = slideNum + 1;
        var objLayout = this.slideLayouts.filter(function (layout) {
            return layout.name == inMasterName;
        })[0];
        // A: Add this SLIDE to PRESENTATION, Add default values as well
        this.slides[slideNum] = {
            name: 'Slide ' + pageNum,
            numb: pageNum,
            data: [],
            rels: [],
            slideNumberObj: null,
            layoutName: inMasterName || '[ default ]',
            layoutObj: objLayout,
        };
        // ==========================================================================
        // PUBLIC METHODS:
        // ==========================================================================
        slideObj.getPageNumber = function () {
            return pageNum;
        };
        slideObj.slideNumber = function (inObj) {
            if (inObj) {
                // A:
                _this.slides[slideNum].slideNumberObj = inObj;
                // B: Add slideNumber to slideMaster1.xml
                if (!_this.masterSlide.slideNumberObj)
                    _this.masterSlide.slideNumberObj = inObj;
                // C: Add slideNumber to `BLANK` (default) layout
                if (!_this.slideLayouts[0].slideNumberObj)
                    _this.slideLayouts[0].slideNumberObj = inObj;
            }
            else {
                return _this.slides[slideNum].slideNumberObj;
            }
        };
        /**
         * Generate the chart based on input data.
         * @see OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
         *
         * @param {object} renderType should belong to: 'column', 'pie'
         * @param {object} data a JSON object with follow the following format
         * {
         *   title: 'eSurvey chart',
         *   data: [
         *		{
         *			name: 'Income',
         *			labels: ['2005', '2006', '2007', '2008', '2009'],
         *			values: [23.5, 26.2, 30.1, 29.5, 24.6]
         *		},
         *		{
         *			name: 'Expense',
         *			labels: ['2005', '2006', '2007', '2008', '2009'],
         *			values: [18.1, 22.8, 23.9, 25.1, 25]
         *		}
         *	 ]
         * }
         */
        slideObj.addChart = function (type, data, opt) {
            genXml.gObjPptxGenerators.addChartDefinition(type, data, opt, _this.slides[slideNum]);
            return _this;
        };
        /**
         * NOTE: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
         */
        slideObj.addImage = function (objImage) {
            genXml.gObjPptxGenerators.addImageDefinition(objImage, this.slides[slideNum]);
            return this;
        };
        slideObj.addMedia = function (opt) {
            var intRels = 1;
            var intImages = ++this.imageCounter;
            var intPosX = opt.x || 0;
            var intPosY = opt.y || 0;
            var intSizeX = opt.w || 2;
            var intSizeY = opt.h || 2;
            var strData = opt.data || '';
            var strLink = opt.link || '';
            var strPath = opt.path || '';
            var strType = opt.type || 'audio';
            var strExtn = 'mp3';
            // STEP 1: REALITY-CHECK
            if (!strPath && !strData && strType != 'online') {
                console.error("ERROR: `addMedia()` requires either 'data' or 'path' values!");
                return null;
            }
            else if (strData && strData.toLowerCase().indexOf('base64,') == -1) {
                console.error("ERROR: Media `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')");
                return null;
            }
            // Online Video: requires `link`
            if (strType == 'online' && !strLink) {
                console.error('addMedia() error: online videos require `link` value');
                return null;
            }
            // STEP 2: Set vars for this Slide
            var slideObjNum = this.slides[slideNum].data.length;
            var slideObjRels = this.slides[slideNum].rels;
            strType = strData ? strData.split(';')[0].split('/')[0] : strType;
            strExtn = strData ? strData.split(';')[0].split('/')[1] : strPath.split('.').pop();
            this.slides[slideNum].data[slideObjNum] = {};
            this.slides[slideNum].data[slideObjNum].type = 'media';
            this.slides[slideNum].data[slideObjNum].mtype = strType;
            this.slides[slideNum].data[slideObjNum].media = strPath || 'preencoded.mov';
            // STEP 3: Set media properties & options
            this.slides[slideNum].data[slideObjNum].options = {};
            this.slides[slideNum].data[slideObjNum].options.x = intPosX;
            this.slides[slideNum].data[slideObjNum].options.y = intPosY;
            this.slides[slideNum].data[slideObjNum].options.cx = intSizeX;
            this.slides[slideNum].data[slideObjNum].options.cy = intSizeY;
            // STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
            // NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
            this.slides.forEach(function (slide) {
                intRels += slide.rels.length;
            });
            if (strType == 'online') {
                // Add video
                slideObjRels.push({
                    path: strPath || 'preencoded' + strExtn,
                    data: 'dummy',
                    type: 'online',
                    extn: strExtn,
                    rId: intRels + 1,
                    Target: strLink,
                });
                this.slides[slideNum].data[slideObjNum].mediaRid = slideObjRels[slideObjRels.length - 1].rId;
                // Add preview/overlay image
                slideObjRels.push({
                    path: 'preencoded.png',
                    data: IMG_PLAYBTN,
                    type: 'image/png',
                    extn: 'png',
                    rId: intRels + 2,
                    Target: '../media/image' + intRels + '.png',
                });
            }
            else {
                // Audio/Video files consume *TWO* rId's:
                // <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
                // <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
                slideObjRels.push({
                    path: strPath || 'preencoded' + strExtn,
                    type: strType + '/' + strExtn,
                    extn: strExtn,
                    data: strData || '',
                    rId: intRels + 0,
                    Target: '../media/media' + intImages + '.' + strExtn,
                });
                this.slides[slideNum].data[slideObjNum].mediaRid = slideObjRels[slideObjRels.length - 1].rId;
                slideObjRels.push({
                    path: strPath || 'preencoded' + strExtn,
                    type: strType + '/' + strExtn,
                    extn: strExtn,
                    data: strData || '',
                    rId: intRels + 1,
                    Target: '../media/media' + intImages + '.' + strExtn,
                });
                // Add preview/overlay image
                slideObjRels.push({
                    data: IMG_PLAYBTN,
                    path: 'preencoded.png',
                    type: 'image/png',
                    extn: 'png',
                    rId: intRels + 2,
                    Target: '../media/image' + intImages + '.png',
                });
            }
            // LAST: Return this Slide
            return this;
        };
        slideObj.addNotes = function (notes, options) {
            genXml.gObjPptxGenerators.addNotesDefinition(notes, options, this.slides[slideNum]);
            return this;
        };
        slideObj.addShape = function (shape, opt) {
            genXml.gObjPptxGenerators.addShapeDefinition(shape, opt, this.slides[slideNum]);
            return this;
        };
        // RECURSIVE: (sometimes)
        // FUTURE: slideObj.addTable = function(arrTabRows, inOpt){
        // FIXME: Move to genXml.gObjPptxGenerators (as every other object uses a generator #consistency)
        // TODO: dont forget to update the "this.color" refs below to "target.slide.color"!!!
        slideObj.addTable = function (arrTabRows, inOpt) {
            var opt = inOpt && typeof inOpt === 'object' ? inOpt : {};
            // STEP 1: REALITY-CHECK
            if (arrTabRows == null || arrTabRows.length == 0 || !Array.isArray(arrTabRows)) {
                try {
                    console.warn('[warn] addTable: Array expected! USAGE: slide.addTable( [rows], {options} );');
                }
                catch (ex) { }
                return null;
            }
            // STEP 2: Row setup: Handle case where user passed in a simple 1-row array. EX: `["cell 1", "cell 2"]`
            //var arrRows = jQuery.extend(true,[],arrTabRows);
            //if ( !Array.isArray(arrRows[0]) ) arrRows = [ jQuery.extend(true,[],arrTabRows) ];
            var arrRows = arrTabRows;
            if (!Array.isArray(arrRows[0]))
                arrRows = [arrTabRows];
            // STEP 3: Set options
            opt.x = getSmartParseNumber(opt.x || (opt.x == 0 ? 0 : EMU / 2), 'X', objLayout);
            opt.y = getSmartParseNumber(opt.y || (opt.y == 0 ? 0 : EMU), 'Y', objLayout);
            opt.cy = opt.h || opt.cy; // NOTE: Dont set default `cy` - leaving it null triggers auto-rowH in `makeXMLSlide()`
            if (opt.cy)
                opt.cy = getSmartParseNumber(opt.cy, 'Y', objLayout);
            opt.h = opt.cy;
            opt.autoPage = opt.autoPage == false ? false : true;
            opt.fontSize = opt.fontSize || DEF_FONT_SIZE;
            opt.lineWeight = typeof opt.lineWeight !== 'undefined' && !isNaN(Number(opt.lineWeight)) ? Number(opt.lineWeight) : 0;
            opt.margin = opt.margin == 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_PT;
            if (!isNaN(opt.margin))
                opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)];
            if (opt.lineWeight > 1)
                opt.lineWeight = 1;
            else if (opt.lineWeight < -1)
                opt.lineWeight = -1;
            // Set default color if needed (table option > inherit from Slide > default to black)
            if (!opt.color)
                opt.color = opt.color || DEF_FONT_COLOR;
            // Set/Calc table width
            // Get slide margins - start with default values, then adjust if master or slide margins exist
            var arrTableMargin = DEF_SLIDE_MARGIN_IN;
            // Case 1: Master margins
            if (objLayout && typeof objLayout.margin !== 'undefined') {
                if (Array.isArray(objLayout.margin))
                    arrTableMargin = objLayout.margin;
                else if (!isNaN(Number(objLayout.margin)))
                    arrTableMargin = [Number(objLayout.margin), Number(objLayout.margin), Number(objLayout.margin), Number(objLayout.margin)];
            }
            // Case 2: Table margins
            /* FIXME: add `margin` option to slide options
            else if ( slideObj.margin ) {
                if ( Array.isArray(slideObj.margin) ) arrTableMargin = slideObj.margin;
                else if ( !isNaN(Number(slideObj.margin)) ) arrTableMargin = [Number(slideObj.margin), Number(slideObj.margin), Number(slideObj.margin), Number(slideObj.margin)];
            }
        */
            // Calc table width depending upon what data we have - several scenarios exist (including bad data, eg: colW doesnt match col count)
            if (opt.w || opt.cx) {
                opt.cx = getSmartParseNumber(opt.w || opt.cx, 'X', objLayout);
                opt.w = opt.cx;
            }
            else if (opt.colW) {
                if (typeof opt.colW === 'string' || typeof opt.colW === 'number') {
                    opt.cx = Math.floor(Number(opt.colW) * arrRows[0].length);
                    opt.w = opt.cx;
                }
                else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length != arrRows[0].length) {
                    console.warn('addTable: colW.length != data.length! Defaulting to evenly distributed col widths.');
                    var numColWidth = Math.floor((_this.pptLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3]) / arrRows[0].length);
                    opt.colW = [];
                    for (var idx = 0; idx < arrRows[0].length; idx++) {
                        opt.colW.push(numColWidth);
                    }
                    opt.cx = Math.floor(numColWidth * arrRows[0].length);
                    opt.w = opt.cx;
                }
            }
            else {
                var numTabWidth = _this.pptLayout.width / EMU - arrTableMargin[1] - arrTableMargin[3];
                opt.cx = Math.floor(numTabWidth);
                opt.w = opt.cx;
            }
            // STEP 4: Convert units to EMU now (we use different logic in makeSlide->table - smartCalc is not used)
            if (opt.x < 20)
                opt.x = inch2Emu(opt.x);
            if (opt.y < 20)
                opt.y = inch2Emu(opt.y);
            if (opt.cx < 20)
                opt.cx = inch2Emu(opt.cx);
            if (opt.cy && opt.cy < 20)
                opt.cy = inch2Emu(opt.cy);
            // STEP 5: Check for fine-grained formatting, disable auto-page when found
            // Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
            // Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
            arrRows.forEach(function (row, rIdx) {
                row.forEach(function (cell, cIdx) {
                    if (cell && cell.text && Array.isArray(cell.text))
                        opt.autoPage = false;
                });
            });
            // STEP 6: Create hyperlink rels
            genXml.createHyperlinkRels(_this.slides, arrRows, _this.slides[slideNum].rels);
            // STEP 7: Auto-Paging: (via {options} and used internally)
            // (used internally by `addSlidesForTable()` to not engage recursion - we've already paged the table data, just add this one)
            if (opt && opt.autoPage == false) {
                // Add data (NOTE: Use `extend` to avoid mutation)
                _this.slides[slideNum].data[_this.slides[slideNum].data.length] = {
                    type: SLIDE_OBJECT_TYPES.table,
                    arrTabRows: arrRows,
                    options: jQuery.extend(true, {}, opt),
                };
            }
            else {
                // Loop over rows and create 1-N tables as needed (ISSUE#21)
                _this.getSlidesForTableRows(arrRows, opt).forEach(function (arrRows, idx) {
                    // A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
                    var currSlide = !this.slides[slideNum + idx] ? this.addNewSlide(inMasterName) : this.slides[slideNum + idx].slide;
                    // B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
                    if (idx > 0)
                        opt.y = inch2Emu(opt.newPageStartY || arrTableMargin[0]);
                    // C: Add this table to new Slide
                    opt.autoPage = false;
                    currSlide.addTable(arrRows, jQuery.extend(true, {}, opt));
                });
            }
            // LAST: Return this Slide
            return _this;
        };
        slideObj.addText = function (text, options) {
            genXml.gObjPptxGenerators.addTextDefinition(text, options, this.slides[slideNum], false);
            return this;
        };
        // ==========================================================================
        // POST-METHODS:
        // ==========================================================================
        // NOTE: Slide Numbers: In order for Slide Numbers to work normally, they need to be in all 3 files: master/layout/slide
        // `defineSlideMaster` and `slideObj.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
        // so, lastly, add to the Slide now.
        if (objLayout && objLayout.slideNumberObj && !slideObj.slideNumber())
            this.slides[slideNum].slideNumberObj = objLayout.slideNumberObj;
        // LAST: Return this Slide
        return slideObj;
    };
    /**
     * Adds a new slide master [layout] to the presentation.
     * @param {ILayout} inObjMasterDef - layout definition
     * @return {ISlide} this slide
     */
    PptxGenJS.prototype.defineSlideMaster = function (inObjMasterDef) {
        if (!inObjMasterDef.title) {
            throw Error('defineSlideMaster() object argument requires a `title` value.');
        }
        var objLayout = {
            name: inObjMasterDef.title,
            slide: null,
            data: [],
            rels: [],
            margin: inObjMasterDef.margin || DEF_SLIDE_MARGIN_IN,
            slideNumberObj: inObjMasterDef.slideNumber || null,
        };
        // STEP 1: Create the Slide Master/Layout
        genXml.gObjPptxGenerators.createSlideObject(inObjMasterDef, objLayout);
        // STEP 2: Add it to layout defs
        this.slideLayouts.push(objLayout);
        // STEP 3: Add slideNumber to master slide (if any)
        if (objLayout.slideNumberObj && !this.masterSlide.slideNumberObj)
            this.masterSlide.slideNumberObj = objLayout.slideNumberObj;
        // LAST:
        return this;
    };
    /**
     * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
     * "Auto-Paging is the future!" --Elon Musk
     *
     * @param {string} tabEleId - The HTML Element ID of the table
     * @param {array} inOpts - An array of options (e.g.: tabsize)
     */
    PptxGenJS.prototype.addSlidesForTable = function (tabEleId, inOpts) {
        var api = this;
        var opts = inOpts || {};
        var arrObjTabHeadRows = [], arrObjTabBodyRows = [], arrObjTabFootRows = [];
        var arrObjSlides = [], arrRows = [], arrColW = [], arrTabColW = [];
        var intTabW = 0, emuTabCurrH = 0;
        // REALITY-CHECK:
        if (jQuery('#' + tabEleId).length == 0) {
            console.error('Table "' + tabEleId + '" does not exist!');
            return;
        }
        var arrInchMargins = [0.5, 0.5, 0.5, 0.5]; // TRBL-style
        opts.margin = opts.margin || opts.margin == 0 ? opts.margin : 0.5;
        if (opts.master && typeof opts.master === 'string') {
            var objLayout = this.slideLayouts.filter(function (layout) {
                return layout.name == opts.master;
            })[0];
            if (objLayout && objLayout.margin) {
                if (Array.isArray(objLayout.margin))
                    arrInchMargins = objLayout.margin;
                else if (!isNaN(objLayout.margin))
                    arrInchMargins = [objLayout.margin, objLayout.margin, objLayout.margin, objLayout.margin];
                opts.margin = arrInchMargins;
            }
        }
        else if (opts && opts.margin) {
            if (Array.isArray(opts.margin))
                arrInchMargins = opts.margin;
            else if (!isNaN(opts.margin))
                arrInchMargins = [opts.margin, opts.margin, opts.margin, opts.margin];
        }
        var emuSlideTabW = opts.w ? inch2Emu(opts.w) : this.pptLayout.width - inch2Emu(arrInchMargins[1] + arrInchMargins[3]);
        var emuSlideTabH = opts.h ? inch2Emu(opts.h) : this.pptLayout.height - inch2Emu(arrInchMargins[0] + arrInchMargins[2]);
        // STEP 1: Grab table col widths
        jQuery.each(['thead', 'tbody', 'tfoot'], function (i, val) {
            if (jQuery('#' + tabEleId + ' > ' + val + ' > tr').length > 0) {
                jQuery('#' + tabEleId + ' > ' + val + ' > tr:first-child')
                    .find('> th, > td')
                    .each(function (i, cell) {
                    // FIXME: This is a hack - guessing at col widths when colspan
                    if (jQuery(this).attr('colspan')) {
                        for (var idx = 0; idx < Number(jQuery(this).attr('colspan')); idx++) {
                            arrTabColW.push(Math.round(jQuery(this).outerWidth() / Number(jQuery(this).attr('colspan'))));
                        }
                    }
                    else {
                        arrTabColW.push(jQuery(this).outerWidth());
                    }
                });
                return false; // break out of .each loop
            }
        });
        jQuery.each(arrTabColW, function (i, colW) {
            intTabW += colW;
        });
        // STEP 2: Calc/Set column widths by using same column width percent from HTML table
        jQuery.each(arrTabColW, function (i, colW) {
            var intCalcWidth = Number(((emuSlideTabW * ((colW / intTabW) * 100)) / 100 / EMU).toFixed(2));
            var intMinWidth = jQuery('#' + tabEleId + ' thead tr:first-child th:nth-child(' + (i + 1) + ')').data('pptx-min-width');
            var intSetWidth = jQuery('#' + tabEleId + ' thead tr:first-child th:nth-child(' + (i + 1) + ')').data('pptx-width');
            arrColW.push(intSetWidth ? intSetWidth : intMinWidth > intCalcWidth ? intMinWidth : intCalcWidth);
        });
        // STEP 3: Iterate over each table element and create data arrays (text and opts)
        // NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
        jQuery.each(['thead', 'tbody', 'tfoot'], function (i, val) {
            jQuery('#' + tabEleId + ' > ' + val + ' > tr').each(function (i, row) {
                var arrObjTabCells = [];
                jQuery(row)
                    .find('> th, > td')
                    .each(function (i, cell) {
                    // A: Get RGB text/bkgd colors
                    var arrRGB1 = [];
                    var arrRGB2 = [];
                    arrRGB1 = jQuery(cell)
                        .css('color')
                        .replace(/\s+/gi, '')
                        .replace('rgba(', '')
                        .replace('rgb(', '')
                        .replace(')', '')
                        .split(',');
                    arrRGB2 = jQuery(cell)
                        .css('background-color')
                        .replace(/\s+/gi, '')
                        .replace('rgba(', '')
                        .replace('rgb(', '')
                        .replace(')', '')
                        .split(',');
                    // ISSUE#57: jQuery default is this rgba value of below giving unstyled tables a black bkgd, so use white instead (FYI: if cell has `background:#000000` jQuery returns 'rgb(0, 0, 0)', so this soln is pretty solid)
                    if (jQuery(cell).css('background-color') == 'rgba(0, 0, 0, 0)' || jQuery(cell).css('background-color') == 'transparent')
                        arrRGB2 = [255, 255, 255];
                    // B: Create option object
                    var objOpts = {
                        fontSize: jQuery(cell)
                            .css('font-size')
                            .replace(/[a-z]/gi, ''),
                        bold: jQuery(cell).css('font-weight') == 'bold' || Number(jQuery(cell).css('font-weight')) >= 500 ? true : false,
                        color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
                        fill: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2])),
                        align: null,
                        border: null,
                        margin: null,
                        colspan: null,
                        rowspan: null,
                        valign: null,
                    };
                    if (['left', 'center', 'right', 'start', 'end'].indexOf(jQuery(cell).css('text-align')) > -1)
                        objOpts.align = jQuery(cell)
                            .css('text-align')
                            .replace('start', 'left')
                            .replace('end', 'right');
                    if (['top', 'middle', 'bottom'].indexOf(jQuery(cell).css('vertical-align')) > -1)
                        objOpts.valign = jQuery(cell).css('vertical-align');
                    // C: Add padding [margin] (if any)
                    // NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
                    if (jQuery(cell).css('padding-left')) {
                        objOpts.margin = [];
                        jQuery.each(['padding-top', 'padding-right', 'padding-bottom', 'padding-left'], function (i, val) {
                            objOpts.margin.push(Math.round(Number(jQuery(cell)
                                .css(val)
                                .replace(/\D/gi, ''))));
                        });
                    }
                    // D: Add colspan/rowspan (if any)
                    if (jQuery(cell).attr('colspan'))
                        objOpts.colspan = jQuery(cell).attr('colspan');
                    if (jQuery(cell).attr('rowspan'))
                        objOpts.rowspan = jQuery(cell).attr('rowspan');
                    // E: Add border (if any)
                    if (jQuery(cell).css('border-top-width') ||
                        jQuery(cell).css('border-right-width') ||
                        jQuery(cell).css('border-bottom-width') ||
                        jQuery(cell).css('border-left-width')) {
                        objOpts.border = [];
                        jQuery.each(['top', 'right', 'bottom', 'left'], function (i, val) {
                            var intBorderW = Math.round(Number(jQuery(cell)
                                .css('border-' + val + '-width')
                                .replace('px', '')));
                            var arrRGB = [];
                            arrRGB = jQuery(cell)
                                .css('border-' + val + '-color')
                                .replace(/\s+/gi, '')
                                .replace('rgba(', '')
                                .replace('rgb(', '')
                                .replace(')', '')
                                .split(',');
                            var strBorderC = rgbToHex(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2]));
                            objOpts.border.push({ pt: intBorderW, color: strBorderC });
                        });
                    }
                    // F: Massage cell text so we honor linebreak tag as a line break during line parsing
                    var $cell2 = jQuery(cell).clone();
                    $cell2.html(jQuery(cell)
                        .html()
                        .replace(/<br[^>]*>/gi, '\n'));
                    // LAST: Add cell
                    arrObjTabCells.push({
                        text: jQuery.trim($cell2.text()),
                        opts: objOpts,
                    });
                });
                switch (val) {
                    case 'thead':
                        arrObjTabHeadRows.push(arrObjTabCells);
                        break;
                    case 'tbody':
                        arrObjTabBodyRows.push(arrObjTabCells);
                        break;
                    case 'tfoot':
                        arrObjTabFootRows.push(arrObjTabCells);
                        break;
                    default:
                }
            });
        });
        // STEP 4: NOTE: `margin` is "cell margin (pt)" everywhere else tables are used, so explicitly convert to "slide margin" here
        if (opts.margin) {
            opts.slideMargin = opts.margin;
            delete opts.margin;
        }
        // STEP 5: Break table into Slides as needed
        // Pass head-rows as there is an option to add to each table and the parse func needs this daa to fulfill that option
        opts.arrObjTabHeadRows = arrObjTabHeadRows || '';
        opts.colW = arrColW;
        this.getSlidesForTableRows(arrObjTabHeadRows.concat(arrObjTabBodyRows).concat(arrObjTabFootRows), opts).forEach(function (arrTabRows, idx) {
            // A: Create new Slide
            var newSlide = api.addNewSlide(opts.master || null);
            // B: DESIGN: Reset `y` to `newPageStartY` or margin after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
            if (idx == 0)
                opts.y = opts.y || arrInchMargins[0];
            if (idx > 0)
                opts.y = opts.newPageStartY || arrInchMargins[0];
            if (opts.debug)
                console.log('opts.newPageStartY:' + opts.newPageStartY + ' / arrInchMargins[0]:' + arrInchMargins[0] + ' => opts.y = ' + opts.y);
            // C: Add table to Slide
            newSlide.addTable(arrTabRows, { x: opts.x || arrInchMargins[3], y: opts.y, w: emuSlideTabW / EMU, colW: arrColW, autoPage: false });
            // D: Add any additional objects
            if (opts.addImage)
                newSlide.addImage({ path: opts.addImage.url, x: opts.addImage.x, y: opts.addImage.y, w: opts.addImage.w, h: opts.addImage.h });
            if (opts.addShape)
                newSlide.addShape(opts.addShape.shape, opts.addShape.opts || opts.addShape.options || {});
            if (opts.addTable)
                newSlide.addTable(opts.addTable.rows, opts.addTable.opts || opts.addTable.options || {});
            if (opts.addText)
                newSlide.addText(opts.addText.text, opts.addText.opts || opts.addText.options || {});
        });
    };
    return PptxGenJS;
}());
export default PptxGenJS;
/*
// NodeJS support
if (this.NODEJS) {
    jQuery = null
    fs = null
    https = null
    JSZip = null
    sizeOf = null

    // A: jQuery dependency
    try {
        var jsdom = require('jsdom')
        var dom = new jsdom.JSDOM('<!DOCTYPE html>')
        jQuery = require('jquery')(dom.window)
    } catch (ex) {
        console.error('Unable to load `jquery`!\n' + ex)
        throw 'LIB-MISSING-JQUERY'
    }

    // B: Other dependencies
    try {
        fs = require('fs')
    } catch (ex) {
        console.error('Unable to load `fs`')
        throw 'LIB-MISSING-FS'
    }
    try {
        https = require('https')
    } catch (ex) {
        console.error('Unable to load `https`')
        throw 'LIB-MISSING-HTTPS'
    }
    try {
        JSZip = require('jszip')
    } catch (ex) {
        console.error('Unable to load `jszip`')
        throw 'LIB-MISSING-JSZIP'
    }
    try {
        sizeOf = require('image-size')
    } catch (ex) {
        console.error('Unable to load `image-size`')
        throw 'LIB-MISSING-IMGSIZE'
    }

    // LAST: Export module
    module.exports = PptxGenJS
}
// Angular/React/etc support
else if (APPJS) {
    // A: jQuery dependency
    try {
        jQuery = require('jquery')
    } catch (ex) {
        console.error('Unable to load `jquery`!\n' + ex)
        throw 'LIB-MISSING-JQUERY'
    }

    // B: Other dependencies
    try {
        JSZip = require('jszip')
    } catch (ex) {
        console.error('Unable to load `jszip`')
        throw 'LIB-MISSING-JSZIP'
    }

    // LAST: Export module
    module.exports = PptxGenJS
}
*/
