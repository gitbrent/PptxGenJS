//export function makeXmlSlideLayoutRel(inSlideNum) {
export default function makeXmlSlideLayoutRel() {
  var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
    strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\r\n';
    //?strXml += '  <Relationship Id="rId'+ inSlideNum +'" Target="../slideMasters/slideMaster'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
    //strXml += '  <Relationship Id="rId1" Target="../slideMasters/slideMaster'+ inSlideNum +'.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>';
    strXml += '  <Relationship Id="rId1" Target="../slideMasters/slideMaster1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"/>\r\n';
    strXml += '</Relationships>';
  //
  return strXml;
}
