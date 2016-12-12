export default function makeXmlSlideRel(inSlideNum, gObjPptx) {
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n'
        + '  <Relationship Id="rId1" Target="../slideLayouts/slideLayout'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"/>\r\n';

    // Add any IMAGEs for this Slide
    for ( var idx=0; idx<gObjPptx.slides[inSlideNum-1].rels.length; idx++ ) {
        strXml += '  <Relationship Id="rId'+ gObjPptx.slides[inSlideNum-1].rels[idx].rId +'" Target="'+ gObjPptx.slides[inSlideNum-1].rels[idx].Target +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>\r\n';
    }

    strXml += '</Relationships>';
    //
    return strXml;
}