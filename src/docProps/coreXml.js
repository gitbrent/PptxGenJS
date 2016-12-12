export default function makeXmlCore( gObjPptx ) {
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n\
          <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"\
             xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"\
             xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\
            <dc:title>' + gObjPptx.title + '</dc:title>\
            <dc:creator>PptxGenJS</dc:creator>\
            <cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy>\
            <cp:revision>1</cp:revision>\
            <dcterms:created xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:created>\
            <dcterms:modified xsi:type="dcterms:W3CDTF">' + new Date().toISOString() + '</dcterms:modified>\
          </cp:coreProperties>';
    return strXml;
}
