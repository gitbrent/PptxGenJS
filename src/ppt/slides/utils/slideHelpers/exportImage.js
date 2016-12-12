export default function (idx, slideObj, locationAttr, x , y, cx, cy) {

    let strSlideXml = `<p:pic>
                        <p:nvPicPr>
                          <p:cNvPr id="${(idx + 2)}" name="Object ${(idx + 1)}" descr="${slideObj.image}"/>
                                <p:cNvPicPr>
                                    <a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/>
                                </p:nvPicPr>
                                <p:blipFill>
                                    <a:blip r:embed="rId${slideObj.imageRid}" cstate="print"/><a:stretch><a:fillRect/></a:stretch>
                                </p:blipFill>
                            <p:spPr>
                                <a:xfrm${locationAttr}>
                                    <a:off  x="${x}"  y="${y}"/>
                                    <a:ext cx="${cx}" cy="${cy}"/>
                                </a:xfrm>'
                                <a:prstGeom prst="rect">
                                    <a:avLst/>
                                </a:prstGeom>
                            </p:spPr>
                    </p:pic>`;

    return strSlideXml;

}
