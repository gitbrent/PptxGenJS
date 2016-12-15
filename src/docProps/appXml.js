export default function makeXmlApp(gObjPptx) {
    var strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n\
					<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\
						<TotalTime>0</TotalTime>\
						<Words>0</Words>\
						<Application>Microsoft Office PowerPoint</Application>\
						<PresentationFormat>On-screen Show (4:3)</PresentationFormat>\
						<Paragraphs>0</Paragraphs>\
						<Slides>' + gObjPptx.slides.length + '</Slides>\
						<Notes>0</Notes>\
						<HiddenSlides>0</HiddenSlides>\
						<MMClips>0</MMClips>\
						<ScaleCrop>false</ScaleCrop>\
						<HeadingPairs>\
						  <vt:vector size="4" baseType="variant">\
						    <vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>\
						    <vt:variant><vt:i4>1</vt:i4></vt:variant>\
						    <vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>\
						    <vt:variant><vt:i4>' + gObjPptx.slides.length + '</vt:i4></vt:variant>\
						  </vt:vector>\
						</HeadingPairs>\
						<TitlesOfParts>';
    strXml += '<vt:vector size="' + ( gObjPptx.slides.length + 1 ) + '" baseType="lpstr">';
    strXml += '<vt:lpstr>Office Theme</vt:lpstr>';
    $.each( gObjPptx.slides, function( idx, slideObj ) {
        strXml += '<vt:lpstr>Slide ' + ( idx + 1 ) + '</vt:lpstr>';
    } );
    strXml += ` </vt:vector>\r\n
          </TitlesOfParts>\r\n
          <Company>PptxGenJS</Company>\r\n
          <LinksUpToDate>false</LinksUpToDate>\r\n
          <SharedDoc>false</SharedDoc>\r\n
          <HyperlinksChanged>false</HyperlinksChanged>\r\n
          <AppVersion>15.0000</AppVersion>\r\n
        </Properties>`;
    return strXml;
}
