export default function makeXmlSlideMasterRel( gObjPptx ) {
    // TODO 1.1: create a slideLayout for each SLDIE
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n';
    for ( var idx = 1; idx <= gObjPptx.slides.length; idx++ ) {
        strXml += '  <Relationship Id="rId' + idx + '" Target="../slideLayouts/slideLayout' + idx + '.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"/>\r\n';
    }
    strXml += '  <Relationship Id="rId' + ( gObjPptx.slides.length + 1 ) + '" Target="../theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/>\r\n';
    strXml += '</Relationships>';
    //
    return strXml;
}
