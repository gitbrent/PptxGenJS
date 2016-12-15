import { BASE_SHAPES, LAYOUTS, APP_VER } from './constante';
import { convertImgToDataURLviaCanvas, callbackImgToDataURLDone } from './ppt/slides/utils/helpers';
import Slide from './ppt/slides/slide';
import ExportPptx from './exportToPptx';
import SlideTable from './ppt/slides/table/slideTable';

export default class PptxGenJS {

    constructor() {
        //this.slideNum = 0;
        this.shapes  = (typeof gObjPptxShapes  !== 'undefined') ? gObjPptxShapes  : BASE_SHAPES;
      	this.masters = (typeof gObjPptxMasters !== 'undefined') ? gObjPptxMasters : {};
    }


    /**
     * Gets the version of this library
     */
    getVersion() {
        return APP_VER;
    };

    /**
     * Sets the Presentation's Title
     */
    setTitle( inStrTitle ) {
        Slide.gObjPptx.title = inStrTitle;
    }

    setLayout(inLayout){

        if ( $.inArray( inLayout, Object.keys( LAYOUTS ) ) > -1 ) {
            Slide.gObjPptx.pptLayout = LAYOUTS[ inLayout ];
        } else {
            try {
                console.warn( 'UNKNOWN LAYOUT! Valid values = ' + Object.keys( LAYOUTS ) );
            } catch ( ex ) {}
        }
        return this;
    }

    /**
     * Gets the Presentation's Slide Layout {object}: [screen4x3, screen16x9, widescreen]
     */
    /*    getLayout() {
        return this._layout;
    }*/



    addNewSlide(isGroup) {
        let slide = new Slide(isGroup);
        return slide.addNewSlide();
    }

    addSlidesForTable(tabEleId, inOpts) {
        let slideTable = new SlideTable();
        slideTable.addSlidesForTable(tabEleId, inOpts);
    }

    /**
     * Export the Presentation to an .pptx file
     * @param {string} [inStrExportName] - Filename to use for the export
     */
    save( inStrExportName ) {

        let exportPptx = new ExportPptx();
        exportPptx.save(inStrExportName);
    }

}
