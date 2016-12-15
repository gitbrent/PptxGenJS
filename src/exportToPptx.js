import ContentType from './contentType';
import rels from './_rels/rels';
import appXml from './docProps/appXml';
import coreXml from './docProps/coreXml';
import presentationXmlRels from './ppt/_rels/presentationXmlRels';
import slideLayoutXml from './ppt/slideLayouts/slideLayoutXml';
import slideLayoutRelXml from './ppt/slideLayouts/slideLayoutRelXml';
import slideMasterXml from './ppt/slideMasters/slideMasterXml';
import slideMasterRelXml from './ppt/slideMasters/slideMasterRelXml';
import themeXml from './ppt/theme/themeXml';
import presentationXml from './ppt/presentationXml';
import presPropsXml from './ppt/presPropsXml';
import tableStyleXml from './ppt/tableStyleXml';
import viewPropsXml from './ppt/viewPropsXml';
import slideXml from './ppt/slides/slideXml';
import slideXmlRel from './ppt/slides/_rel/slideXmlRel';
import Slide from 'ppt/slides/slide.js';


export default class ExportPptx {

    constructor() {

        this.gObjPptx = Slide.gObjPptx;
        //this.zip = new JSZip();
        this.zip = new com.sap.powerdesigner.web.galilei.common.util.Zip.create();
        this.intSlideNum = 0;
        this.intRels = 0;
    }

    save(inStrExportName){

        let intRels = 0,
            arrImages = [];

        // STEP 1: Set export title (if any)
        if ( inStrExportName ) this.gObjPptx.fileName = inStrExportName;

        // STEP 2: Total all images (rels) across the Presentation
        // PERF: Only send unique image paths for encoding (encoding func will find and fill ALL matching img paths and fill)
        let slides = this.gObjPptx.slides;
        $.each( slides, ( i, slide ) => {
            $.each( slide.rels, ( i, rel ) => {
                if ( !rel.data && $.inArray( rel.path, arrImages ) == -1 ) {
                    intRels++;
                    this.convertImgToDataURLviaCanvas( rel, this.callbackImgToDataURLDone );
                    arrImages.push( rel.path );
                }
            } );
        } );

        // STEP 3: Export now if there's no images to encode (otherwise, last async imgConvert call above will call exportFile)
        if ( intRels == 0 ) {
            this.doExportPresentation();
        };
    }

    doExportPresentation() {
        this.zip.folder( "_rels" );
        this.zip.folder( "docProps" );
        this.zip.folder( "ppt" ).folder( "_rels" );
        this.zip.folder( "ppt/media" );
        this.zip.folder( "ppt/slideLayouts" ).folder( "_rels" );
        this.zip.folder( "ppt/slideMasters" ).folder( "_rels" );
        this.zip.folder( "ppt/slides" ).folder( "_rels" );
        this.zip.folder( "ppt/theme" );

        this.zip.file( "[Content_Types].xml", ContentType(this.gObjPptx) );
        this.zip.file( "_rels/.rels", rels() );
        this.zip.file( "docProps/app.xml", appXml(this.gObjPptx) );
        this.zip.file( "docProps/core.xml", coreXml(this.gObjPptx) );
        this.zip.file( "ppt/_rels/presentation.xml.rels", presentationXmlRels(this.gObjPptx) );

        // Create a Layout/Master/Rel/Slide file for each SLIDE
        for ( var idx = 0; idx < this.gObjPptx.slides.length; idx++ ) {
            this.intSlideNum++;
            this.zip.file( "ppt/slideLayouts/slideLayout" + this.intSlideNum + ".xml", slideLayoutXml() );
            this.zip.file( "ppt/slideLayouts/_rels/slideLayout" + this.intSlideNum + ".xml.rels", slideLayoutRelXml() );
            this.zip.file( "ppt/slides/slide" + this.intSlideNum + ".xml", slideXml( this.gObjPptx.slides[ idx ], this.gObjPptx ) );
            this.zip.file( "ppt/slides/_rels/slide" + this.intSlideNum + ".xml.rels", slideXmlRel( this.intSlideNum, this.gObjPptx ) );
        }
        this.zip.file( "ppt/slideMasters/slideMaster1.xml", slideMasterXml(this.gObjPptx) );
        this.zip.file( "ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRelXml(this.gObjPptx) );

        // Add all images
        this.addAllImages();

        this.zip.file( "ppt/theme/theme1.xml", themeXml() );
        this.zip.file( "ppt/presentation.xml", presentationXml(this.gObjPptx) );
        this.zip.file( "ppt/presProps.xml", presPropsXml() );
        this.zip.file( "ppt/tableStyles.xml", tableStyleXml() );
        this.zip.file( "ppt/viewProps.xml", viewPropsXml() );

        // =======
        // STEP 3: Push the PPTX file to browser
        // =======
        var strExportName = ( ( this.gObjPptx.fileName.toLowerCase().indexOf( '.ppt' ) > -1 ) ? this.gObjPptx.fileName : this.gObjPptx.fileName + this.gObjPptx.fileExtn );
        this.zip.generateAsync( {
            type: "blob"
        } ).then( function( content ) {
            //saveAs( content, strExportName );
            sap.galilei.ui.common.FileManager.saveAs( content, strExportName );
            Slide.gObjPptx.slides = [];
        });

    }

    addAllImages() {
        for (var idx = 0; idx < this.gObjPptx.slides.length; idx++) {
            for (var idy = 0; idy < this.gObjPptx.slides[idx].rels.length; idy++) {
                var id = this.gObjPptx.slides[idx].rels[idy].rId - 1;
                var data = this.gObjPptx.slides[idx].rels[idy].data;
                // data:image/png;base64
                var header = data.substring(0, data.indexOf(","));
                // NOTE: Trim the leading 'data:image/png;base64,' text as it is not needed (and image wont render correctly with it)
                var content = data.substring(data.indexOf(",") + 1);
                var extn = /data:image\/(\w+)/.exec(header)[1];
                var isBase64 = /base64/.test(header);
                this.zip.file("ppt/media/image" + id + "." + extn, content, {
                    base64: isBase64
                });
            }
        }
    }

    convertImgToDataURLviaCanvas(slideRel) {
        // A: Create
        var self = this;
        var image = new Image();
        // B: Set onload event
        image.onload = function() {
            // First: Check for any errors: This is the best method (try/catch wont work, etc.)
            if (this.width + this.height == 0) {
                this.onerror();
                return;
            }
            var canvas = document.createElement('CANVAS');
            var ctx = canvas.getContext('2d');
            canvas.height = this.height;
            canvas.width = this.width;
            ctx.drawImage(this, 0, 0);
            // Users running on local machine will get the following error:
            // "SecurityError: Failed to execute 'toDataURL' on 'HTMLCanvasElement': Tainted canvases may not be exported."
            // when the canvas.toDataURL call executes below.
            try {
                self.callbackImgToDataURLDone(canvas.toDataURL(slideRel.type), slideRel);
            } catch (ex) {
                this.onerror();
                console.log("NOTE: Browsers wont let you load/convert local images! (search for --allow-file-access-from-files)");
                return;
            }
            canvas = null;
        };
        image.onerror = function() {
            try {
                console.error('[Error] Unable to load image: ' + slideRel.path);
            } catch (ex) {}
            // Return a predefined "Broken image" graphic so the user will see something on the slide
            self.callbackImgToDataURLDone('data:image/png;base64,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', slideRel);
        };
        // C: Load image
        image.src = slideRel.path;
    }

    callbackImgToDataURLDone(inStr, slideRel) {
        var intEmpty = 0;

        // STEP 1: Store base64 data for this image
        slideRel.data = inStr;

        // STEP 2: Call export function once all async processes have completed
        $.each(this.gObjPptx.slides, function(i, slide) {
            $.each(slide.rels, function(i, rel) {
                if (rel.path == slideRel.path) rel.data = inStr;
                if (!rel.data) intEmpty++;
            });
        });

        // STEP 3: Continue export process
        if (intEmpty == 0) this.doExportPresentation();
    }

}
