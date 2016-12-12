export default function makeXmlViewProps() {
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>' +
        '<p:slideViewPr>' +
        '  <p:cSldViewPr>' +
        '    <p:cViewPr varScale="1"><p:scale><a:sx n="64" d="100"/><a:sy n="64" d="100"/></p:scale><p:origin x="-1392" y="-96"/></p:cViewPr>' +
        '    <p:guideLst><p:guide orient="horz" pos="2160"/><p:guide pos="2880"/></p:guideLst>' +
        '  </p:cSldViewPr>' +
        '</p:slideViewPr>' +
        '<p:notesTextViewPr>' +
        '  <p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr>' +
        '</p:notesTextViewPr>' +
        '<p:gridSpacing cx="78028800" cy="78028800"/>' +
        '</p:viewPr>';
    return strXml;
}
