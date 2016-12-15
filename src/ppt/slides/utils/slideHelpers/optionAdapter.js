import { getShapeInfo, getSmartParseNumber } from '../helpers';
import Slide from '../../slide.js'

export default function(slideObj){
    const EMU = 914400;
    let x = 0, y = 0, cx = (EMU*10), cy = 0,
        locationAttr = '',
        shapeType = null;

    if ( slideObj.options.shape  ) shapeType = getShapeInfo( slideObj.options.shape );
    if ( slideObj.options.w  || slideObj.options.w  == 0 ) slideObj.options.cx = slideObj.options.w;
    if ( slideObj.options.h  || slideObj.options.h  == 0 ) slideObj.options.cy = slideObj.options.h;
    if (shapeType && shapeType.name === 'polyline'){
        slideObj.options.x  = 0;
        slideObj.options.y  = 0;
    }else {
        if ( slideObj.options.x  || slideObj.options.x  == 0 )  x = getSmartParseNumber( slideObj.options.x , 'X', Slide.gObjPptx );
        if ( slideObj.options.y  || slideObj.options.y  == 0 )  y = getSmartParseNumber( slideObj.options.y , 'Y', Slide.gObjPptx );
    }
    if ( slideObj.options.cx || slideObj.options.cx == 0 ) cx = getSmartParseNumber( slideObj.options.cx, 'X', Slide.gObjPptx );
    if ( slideObj.options.cy || slideObj.options.cy == 0 ) cy = getSmartParseNumber( slideObj.options.cy, 'Y', Slide.gObjPptx );
    if ( slideObj.options.flipH  ) locationAttr += ' flipH="1"';
    if ( slideObj.options.flipV  ) locationAttr += ' flipV="1"';
    if ( slideObj.options.rotate ) {
        let rotateVal = (slideObj.options.rotate > 360) ? (slideObj.options.rotate - 360) : slideObj.options.rotate;
        rotateVal *= 60000;
        locationAttr += ` rot="${rotateVal}"`;
    }
    return {x, y,cx,cy,shapeType,locationAttr}
}