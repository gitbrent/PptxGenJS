import ExportTable  from './utils/slideHelpers/exportTable';
import ExportText  from './utils/slideHelpers/exportText';
import ExportImage  from './utils/slideHelpers/exportImage';
import OptionAdapter  from './utils/slideHelpers/optionAdapter';
import Group from './group/slideGroup';
import Slide from './slide';


export default function makeXmlSlide(inSlide, gObjPptx) {

    const EMU = 914400, ONEPT= 12700;
    let intTableNum = 1,
    objSlideData = inSlide.data,
    strSlideXml;

    if(inSlide.slide.group){
        var startGroup, endGroup,
        group = new Group(gObjPptx);
        group.generateGroup();
        startGroup = group.groupStart;
        endGroup = group.groupEnd;
    }
    // STEP 1: Start slide XML
    strSlideXml = Slide.header(inSlide);
    (inSlide.slide.group) ? strSlideXml += startGroup: '';

    // STEP 5: Loop over all Slide objects and add them to this slide:
    $.each(objSlideData, function(idx, slideObj){

        // A: Set option vars
        if ( slideObj.options ) {
            var { x, y, cx, cy, shapeType, locationAttr } = OptionAdapter(slideObj, gObjPptx)
        }else{
            var x = 0, y = 0, cx = (EMU*10), cy = 0,
                locationAttr = '',
                shapeType = null;
        }

        // B: Create this particular object on Slide
        switch ( slideObj.type ) {
            case 'table':
                strSlideXml += ExportTable(inSlide, slideObj, intTableNum, x, y, cx, cy);
                break;

            case 'text':
                strSlideXml += ExportText(inSlide, shapeType, slideObj, locationAttr, idx, x, y, cx, cy);
                break;

            case 'image':
                strSlideXml += ExportImage(idx, slideObj, locationAttr, x , y, cx, cy)
                break;
        }
    });
    (inSlide.slide.group) ? strSlideXml += endGroup: '';
    // STEP 6: Close spTree and finalize slide XML
    strSlideXml += Slide.footer();

    // LAST: Return
    return strSlideXml;
}