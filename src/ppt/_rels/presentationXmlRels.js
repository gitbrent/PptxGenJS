export default function makeXmlPresentationRels(gObjPptx) {
  var intRelNum = 0;
  var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

  strXml += '  <Relationship Id="rId1" Target="slideMasters/slideMaster1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
  intRelNum++;

  for ( var idx=1; idx<=gObjPptx.slides.length; idx++ ) {
    intRelNum++;
    strXml += '  <Relationship Id="rId'+ intRelNum +'" Target="slides/slide'+ idx +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>';
  }
  intRelNum++;
  strXml += '  <Relationship Id="rId'+  intRelNum    +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>'
      + '  <Relationship Id="rId'+ (intRelNum+1) +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>'
      + '  <Relationship Id="rId'+ (intRelNum+2) +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
      + '  <Relationship Id="rId'+ (intRelNum+3) +'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>'
      + '</Relationships>';
  return strXml;
}
