export default function makeXmlTableStyles() {
    // SEE: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
    let strXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n
        <a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
    return strXml;
}
