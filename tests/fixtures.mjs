import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
const { A, P, preset, parseXML } = await import('../src/styles.js');
export const styleXML = preset('red');
export const oldID = parseXML(styleXML).documentElement.getAttribute('styleId');
export async function fixture({ styles = true, overrides = false, masters = [] } = {}) {
  const zip = new JSZip();
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="300" r:id="rId2"/></p:sldIdLst></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="${R}/slide" Target="slides/slide2.xml"/></Relationships>`);
  for (const [i, other] of masters.entries()) {
    zip.file(`ppt/slideMasters/slideMaster${i + 1}.xml`, `<p:sldMaster xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree/></p:cSld>${other === null ? '' : `<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400" b="1"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle>${other}</p:otherStyle></p:txStyles>`}<p:extLst><p:ext uri="keep"/></p:extLst></p:sldMaster>`);
  }
  if (styles) zip.file('ppt/tableStyles.xml', `<a:tblStyleLst xmlns:a="${A}" def="${oldID}">${styleXML}</a:tblStyleLst>`);
  for (let i = 1; i <= 2; i++) zip.file(`ppt/slides/slide${i}.xml`, `<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Table ${i}"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${oldID}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="1000000"/></a:tblGrid><a:tr h="100000"><a:tc><a:txBody><a:p><a:r><a:rPr sz="1800" b="1">${overrides ? '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' : ''}</a:rPr><a:t>UNCHANGED TEXT ${i}</a:t></a:r></a:p></a:txBody><a:tcPr marL="1234">${overrides ? '<a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:lnL/>' : ''}</a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/media/image1.bin', new Uint8Array([0, 255, 4, 128, 0]));
  zip.file('customXml/item1.xml', '<keep>untouched</keep>');
  return zip.generateAsync({ type: 'uint8array' });
}
export function plan(extra = {}) {
  return { id: oldID, name: '我的新蓝色', sourceXML: styleXML, patches: { firstRow: { fill: '#0030C2' } }, scope: 'linked', setDefault: true, ...extra };
}
