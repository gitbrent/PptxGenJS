export default function makeXmlContTypes(gObjPptx) {
    const CRLF = '\r\n';

    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        ' <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        ' <Default Extension="xml" ContentType="application/xml"/>' +
        ' <Default Extension="jpeg" ContentType="image/jpeg"/>' +
        ' <Default Extension="png" ContentType="image/png"/>' +
        ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
        ' <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
        ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        ' <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' +
        ' <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
        ' <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' +
        ' <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>';
    $.each( gObjPptx.slides, function( idx, slide ) {
        strXml += '<Override PartName="/ppt/slideMasters/slideMaster' + ( idx + 1 ) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>';
        strXml += '<Override PartName="/ppt/slideLayouts/slideLayout' + ( idx + 1 ) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>';
        strXml += '<Override PartName="/ppt/slides/slide' + ( idx + 1 ) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>';
    } );
    strXml += '</Types>';
    return strXml;
}
