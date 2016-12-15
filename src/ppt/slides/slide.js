import { inch2Emu, genXmlColorSelection } from './utils/helpers';
import { LAYOUTS} from '../../constante';


const EMU = 914400, SLDNUMFLDID = '{F7021451-1387-4CA6-816F-3879F97B5CBC}';

class Slide {

    constructor( isGroup=false ) {
        this.group = isGroup;
        this.slideNum = 0;
        this.slideObjNum = 0;
    }

    static gObjPptx = {
                        title: 'PresePptxGenJS Presentation',
                        fileName: 'Presentation',
                        fileExtn: '.pptx',
                        pptLayout: LAYOUTS['LAYOUT_WIDE'],
                        slides: []
                      };

    hasSlideNumber(inBool) {
        if (inBool) Slide.gObjPptx.slides[this.slideNum].hasSlideNumber = inBool;
        else return Slide.gObjPptx.slides[this.slideNum].hasSlideNumber;
    }

    getPageNumber() {
        return this.slideNum;
    }
    
    addNewSlide(inMaster){
        this.slideNum = Slide.gObjPptx.slides.length;
        let pageNum = (this.slideNum + 1);

        // A: Add this SLIDE to PRESENTATION, Add default values as well
        Slide.gObjPptx.slides[this.slideNum] = {};
        Slide.gObjPptx.slides[this.slideNum].slide = new Slide(this.group);
        Slide.gObjPptx.slides[this.slideNum].name = 'Slide ' + pageNum;
        Slide.gObjPptx.slides[this.slideNum].numb = pageNum;
        Slide.gObjPptx.slides[this.slideNum].data = [];
        Slide.gObjPptx.slides[this.slideNum].rels = [];
        Slide.gObjPptx.slides[this.slideNum].hasSlideNumber = false;

        // C: Add 'Master Slide' attr to Slide if a valid master was provided
        if (inMaster && this.masters) {
            // A: Add images (do this before adding slide bkgd)
            if (inMaster.images && inMaster.images.length > 0) {
                $.each(inMaster.images, (i, image) =>{
                    this.addImage(image.src, inch2Emu(image.x), inch2Emu(image.y), inch2Emu(image.cx), inch2Emu(image.cy), (image.data || ''));
                });
            }

            // B: Add any Slide Background: Image or Fill
            if (inMaster.bkgd && inMaster.bkgd.src) {
                let slideObjRels = Slide.gObjPptx.slides[this.slideNum].rels;
                let strImgExtn = inMaster.bkgd.src.substring(inMaster.bkgd.src.indexOf('.') + 1).toLowerCase();
                if (strImgExtn == 'jpg') strImgExtn = 'jpeg';
                if (strImgExtn == 'gif') strImgExtn = 'png'; // MS-PPT: canvas.toDataURL for gif comes out image/png, and PPT will show "needs repair" unless we do this
                // TODO 1.5: The next few lines are copies from .addImage above. A bad idea thats already bit my once! So of course it's makred as future :)
                var intRels = 1;
                for (var idx = 0; idx < Slide.gObjPptx.slides.length; idx++) {
                    intRels += Slide.gObjPptx.slides[idx].rels.length;
                }
                slideObjRels.push({
                    path: inMaster.bkgd.src,
                    type: 'image/' + strImgExtn,
                    extn: strImgExtn,
                    data: (inMaster.bkgd.data || ''),
                    rId: (intRels + 1),
                    Target: '../media/image' + intRels + '.' + strImgExtn
                });
                slide.bkgdImgRid = slideObjRels[slideObjRels.length - 1].rId;
            } else if (inMaster.bkgd) {
                slide.back = inMaster.bkgd;
            }

            // C: Add shapes
            if (inMaster.shapes && inMaster.shapes.length > 0) {
                $.each(inMaster.shapes, (i, shape) => {
                    // 1: Grab all options (x, y, color, etc.)
                    var objOpts = {};
                    $.each(Object.keys(shape), (i, key) => {
                        if (shape[key] != 'type') objOpts[key] = shape[key];
                    });
                    // 2: Create object using 'type'
                    if (shape.type == 'text') slide.addText(shape.text, objOpts);
                    else if (shape.type == 'line') slide.addShape(this.shapes.LINE, objOpts);
                });
            }

            // D: Slide Number
            if (typeof inMaster.isNumbered !== 'undefined') this.slide.hasSlideNumber(inMaster.isNumbered);
        }
        return this;
    }

    addTable(arrTabRows, inOpt, tabOpt) {

        var opt = (typeof inOpt === 'object') ? inOpt : {};
        if (opt.w) opt.cx = opt.w;
        if (opt.h) opt.cy = opt.h;

        // STEP 1: REALITY-CHECK
        if (arrTabRows == null || arrTabRows.length == 0 || !Array.isArray(arrTabRows)) {
            try {
                console.warn('[warn] addTable: Array expected!');
            } catch (ex) {}
            return null;
        }

        // STEP 2: Grab Slide object count
        this.slideObjNum = Slide.gObjPptx.slides[this.slideNum].data.length;

        // STEP 3: Set default options if needed
        if (typeof opt.x === 'undefined') opt.x = (EMU / 2);
        if (typeof opt.y === 'undefined') opt.y = EMU;
        if (typeof opt.cx === 'undefined') opt.cx = (Slide.gObjPptx.pptLayout.width - (EMU / 2));
        // Dont do this for cy - leaving it null triggers auto-rowH in makeXMLSlide function

        // STEP 4: We use different logic in makeSlide (smartCalc is not used), so convert to EMU now
        if (opt.x < 20) opt.x = inch2Emu(opt.x);
        if (opt.y < 20) opt.y = inch2Emu(opt.y);
        if (opt.w < 20) opt.w = inch2Emu(opt.w);
        if (opt.h < 20) opt.h = inch2Emu(opt.h);
        if (opt.cx < 20) opt.cx = inch2Emu(opt.cx);
        if (opt.cy && opt.cy < 20) opt.cy = inch2Emu(opt.cy);
        //
        if (tabOpt && Array.isArray(tabOpt.colW)) {
            $.each(tabOpt.colW, function(i, colW) {
                if (colW < 20) tabOpt.colW[i] = inch2Emu(colW);
            });
        }

        // Handle case where user passed in a simple array
        var arrTemp = $.extend(true, [], arrTabRows);
        if (!Array.isArray(arrTemp[0])) arrTemp = [$.extend(true, [], arrTabRows)];

        // STEP 5: Add data
        // NOTE: Use extend to avoid mutation
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum] = {
            type: 'table',
            arrTabRows: arrTemp,
            options: $.extend(true, {}, opt),
            objTabOpts: ($.extend(true, {}, tabOpt) || {})
        };

        // LAST: Return this Slide object
        return this;
    }

    addText(text, opt) {
        // STEP 1: Grab Slide object count
        this.slideObjNum = Slide.gObjPptx.slides[this.slideNum].data.length;

        // ROBUST: Convert attr values that will likely be passed by users to valid OOXML values
        if (opt.valign) opt.valign = opt.valign.toLowerCase().replace(/^c.*/i, 'ctr').replace(/^m.*/i, 'ctr').replace(/^t.*/i, 't').replace(/^b.*/i, 'b');
        if (opt.align) opt.align = opt.align.toLowerCase().replace(/^c.*/i, 'center').replace(/^m.*/i, 'center').replace(/^l.*/i, 'left').replace(/^r.*/i, 'right');

        // STEP 2: Set props
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum] = {};
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].type = 'text';
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].text = text;
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options = (typeof opt === 'object') ? opt : {};
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp = jQuery.extend({}, opt.bodyProp);
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.autoFit = (opt.autoFit || false); // If true, shape will collapse to text size (Fit To Shape)
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.anchor = (opt.valign || 'ctr'); // VALS: [t,ctr,b]
        if ((opt.inset && !isNaN(Number(opt.inset))) || opt.inset == 0) {
            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.lIns = inch2Emu(opt.inset);
            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.rIns = inch2Emu(opt.inset);
            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.tIns = inch2Emu(opt.inset);
            Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.bodyProp.bIns = inch2Emu(opt.inset);
        }

        // LAST: Return
        return this;
    }

    addShape(shape, opt) {
        // STEP 1: Grab Slide object count
        this.slideObjNum = Slide.gObjPptx.slides[this.slideNum].data.length;

        // STEP 2: Set props
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum] = {};
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].type = 'text';
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options = (typeof opt == 'object') ? opt : {};
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.shape = shape;

        return this;
    }

    addImage(strImagePath, intPosX, intPosY, intSizeX, intSizeY, strImageData, strImgData) {
        var intRels = 1;

        // FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
        // TODO: FUTURE: DEPRECATED: Only allow object param in 1.5 or 2.0
        if ( typeof strImagePath === 'object' ) {
            intPosX = (strImagePath.x || 0);
            intPosY = (strImagePath.y || 0);
            intSizeX = (strImagePath.cx || strImagePath.w || 0);
            intSizeY = (strImagePath.cy || strImagePath.h || 0);
            strImageData = (strImagePath.data || '');
            strImagePath = (strImagePath.path || ''); // This line must be last as were about to ovewrite ourself!
        }
        // REALITY-CHECK:
        if (!strImagePath && !strImgData) {
            try {
                console.error("ERROR: Image can't be empty");
            } catch (ex) {}
            return null;
        }

        // STEP 1: Set vars for this Slide
        this.slideObjNum = Slide.gObjPptx.slides[this.slideNum].data.length;
        var slideObjRels = Slide.gObjPptx.slides[this.slideNum].rels;
        var strImgExtn = 'png'; // Every image is encoded via canvas>base64, so they all come out as png (use of another extn will cause "needs repair" dialog on open in PPT)

        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum] = {};
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].type = 'image';
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].image = strImagePath;

        // STEP 2: Set image properties & options
        // TODO 1.1: Measure actual image when no intSizeX/intSizeY params passed
        // ....: This is an async process: we need to make getSizeFromImage use callback, then set H/W...
        // if ( !intSizeX || !intSizeY ) { var imgObj = getSizeFromImage(strImagePath);
        var imgObj = {
            width: 1,
            height: 1
        };
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options = {};
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.x = (intPosX || 0);
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.y = (intPosY || 0);
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.cx = (intSizeX || imgObj.width);
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].options.cy = (intSizeY || imgObj.height);

        // STEP 3: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
        // NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
        $.each(Slide.gObjPptx.slides, function(i, slide) {
            intRels += slide.rels.length;
        });
        slideObjRels.push({
            path: strImagePath,
            type: 'image/' + strImgExtn,
            extn: strImgExtn,
            data: (strImgData || ''),
            rId: (intRels + 1),
            Target: '../media/image' + intRels + '.' + strImgExtn
        });
        Slide.gObjPptx.slides[this.slideNum].data[this.slideObjNum].imageRid = slideObjRels[slideObjRels.length - 1].rId;

        // LAST: Return this Slide
        return this;
    }

    static header(inSlide){

        let strSlideXml, aStr=[], propertBg = [], propertySpTree=[], propertySp=[];

        let head =`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="${inSlide.name}">`;
        aStr.push(head);

        // STEP 2: Add background color or background image (if any)
        // A: Background color
        /*if ( inSlide.slide.back ) strSlideXml += genXmlColorSelection(false, inSlide.slide.back);
        // B: Add background image (using Strech) (if any)
        if ( inSlide.slide.bkgdImgRid ) {
            // TODO 1.0: We should be doing this in the slideLayout...
            strSlideXml += `<p:bg>
                            <p:bgPr><a:blipFill dpi="0" rotWithShape="1">
                                <a:blip r:embed="rId${inSlide.slide.bkgdImgRid}"><a:lum/></a:blip>
                                <a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
                                <a:effectLst/></p:bgPr>
                         </p:bg>`;
        } */

        if ( inSlide.slide.back ) aStrSlideXml.push(genXmlColorSelection(false, inSlide.slide.back));
        // B: Add background image (using Strech) (if any)
        if ( inSlide.slide.bkgdImgRid ) {
            // TODO 1.0: We should be doing this in the slideLayout...
            propertBg = [`<p:bg>`,
                `<p:bgPr><a:blipFill dpi="0" rotWithShape="1">`,
                `<a:blip r:embed="rId${inSlide.slide.bkgdImgRid}"><a:lum/></a:blip>`,
                `<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
                `<a:effectLst/></p:bgPr>`,
                `</p:bg>`
            ];
            aStr.push(propertBg.join(''))
        }
        propertySpTree = [`<p:spTree>`,
            `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`,
            `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>`,
            `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`];

            aStr.push(propertySpTree.join(''));

// STEP 4: Add slide numbers if selected
        // TODO 1.0: Fixed location sucks! Place near bottom corner using slide.size !!!
        if ( inSlide.hasSlideNumber ) {
            propertySp = [`<p:sp>`,
                              `<p:nvSpPr>`,
                              `<p:cNvPr id="25" name="Shape 25"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr></p:nvSpPr>`,
                              `<p:spPr>`,
                                `<a:xfrm><a:off x="${(EMU*0.3)}" y="${(EMU*5.2)}"/><a:ext cx="400000" cy="300000"/></a:xfrm>`,
                                `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`,
                                `<a:extLst>`,
                                `<a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}">`,
                                  `<ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext>`,
                                `</a:extLst>`,
                              `</p:spPr>`,
                              `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr/><a:fld id="${SLDNUMFLDID}" type="slidenum"/></a:p></p:txBody>`,
                              `</p:sp>`];
            aStr.push(propertySp.join(''))

        }
        strSlideXml = aStr.join('');
        return strSlideXml;
    }

    static footer(){
        let footer= [
            `</p:spTree>`,
                `</p:cSld>`,
                `<p:clrMapOvr>`,
                    `<a:masterClrMapping/>`,
                `</p:clrMapOvr>`,
            `</p:sld>`
            ];
        return footer.join('');
    }

}

export default Slide