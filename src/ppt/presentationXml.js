export default function makeXmlPresentation(gObjPptx) {
  var intCurPos = 0;
  var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">\r\n';

  // STEP 1: Build SLIDE master list
  strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>\r\n';
  strXml += '<p:sldIdLst>\r\n';
  for ( var idx=0; idx<gObjPptx.slides.length; idx++ ) {
    strXml += '<p:sldId id="' + (idx + 256) + '" r:id="rId' + (idx + 2) + '"/>\r\n';
  }
  strXml += '</p:sldIdLst>\r\n';

  // STEP 2: Build SLIDE text styles
  strXml += '<p:sldSz cx="'+ gObjPptx.pptLayout.width +'" cy="'+ gObjPptx.pptLayout.height +'" type="'+ gObjPptx.pptLayout.name +'"/>\r\n'
      + '<p:notesSz cx="'+ gObjPptx.pptLayout.height +'" cy="' + gObjPptx.pptLayout.width + '"/>'
      + '<p:defaultTextStyle>';
      + '  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>';
  for ( var idx=1; idx<10; idx++ ) {
    strXml += '  <a:lvl' + idx + 'pPr marL="' + intCurPos + '" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
        + '    <a:defRPr sz="1800" kern="1200">'
        + '      <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        + '      <a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>'
        + '    </a:defRPr>'
        + '  </a:lvl' + idx + 'pPr>';
    intCurPos += 457200;
  }
  strXml += '</p:defaultTextStyle>\r\n';

  strXml += '<p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext></p:extLst>\r\n'
      + '</p:presentation>';
  //
  return strXml;
}
