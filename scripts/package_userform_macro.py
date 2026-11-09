#!/usr/bin/env python3
"""Package only the native editor/palette UserForms; slides contain static help."""
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT/'macro'))
sys.path.insert(0, str(ROOT/'scripts'))
from userform import build_form, definitions, event_code
from userform_color_picker import build_form as build_palette, event_code as palette_events, geometry as palette_geometry
from package_macro import validate_project, shape, P, A, CT, REL

VERSION = json.loads((ROOT/'package.json').read_text())['version']
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def entry_code():
    source = (ROOT/'macro/userform-code.js').read_text()
    source = source.replace('function TsOpenUserForm_Click() { return OpenTableStyleUserForm(); }', '')
    source = source.replace('__TABLE_STYLES_USERFORM_GEOMETRY__', json.dumps(definitions(), ensure_ascii=False))
    source = source.replace('__TABLE_STYLES_USERFORM_PICKER_GEOMETRY__', json.dumps(palette_geometry()))
    return ('function OpenTableStyleEditor() { return OpenTableStyleUserForm(); }\n'
            'function TableStylesVersion() { return '+json.dumps(VERSION)+'; }\n'+source+'\n'+
            event_code(definitions())+'\n'+palette_events()+'''
function DiagnoseTableStyleEditor() {
  var report = "UserForm 宏版 " + TableStylesVersion();
  try {
    var editor = TableStylesFormEditor();
    var platform = editor.host.app.TableStylePlatform;
    report += "\\n文件接口：" + platform.fileBackend + "\\nUUID：" + platform.uuidBackend;
  } catch (e) { report += "\\n" + (e.message || String(e)); }
  if (typeof MsgBox === "function") MsgBox(report, 0, "表格样式编辑器诊断");
  return report;
}
''')


def build_project(bundle=''):
    project = ET.Element('document', version='2.0')
    ET.SubElement(project, 'name').text = 'TableStylesUserForm'
    ET.SubElement(project, 'property', desc='', lock='false', password='')
    ET.SubElement(project, 'activemodule').text = '1'
    source = entry_code()
    for name, number, code in [('TableStyleEditor', '1', source), ('TableStyleRuntime', '2',
            'function TableStylesNativeRuntime() {\nif (!TableStylesNativeRuntime.cached) {\n'+bundle+
            '\nTableStylesNativeRuntime.cached = NativeUserFormTableStyles;\n}\nreturn TableStylesNativeRuntime.cached;\n}\n')]:
        module = ET.SubElement(project, 'codemodule', name=name, id=number)
        ET.SubElement(module, 'window', cursorpos=str(source.index('{')+1 if number == '1' else 0), actived='true' if number == '1' else 'false', visible='false')
        ET.SubElement(module, 'codetext').text = code
    for name, number, form in [('TableStyleForm', '3', build_form()[0]), ('TableStyleColorForm', '4', build_palette()[0])]:
        dialog = ET.SubElement(project, 'dlgmodule', name=name, id=number)
        ET.SubElement(dialog, 'window', visible='false'); dialog.append(form)
    ET.SubElement(project, 'functionsdata')
    validate_project(project)
    return project


def build_entries(bundle=''):
    with ZipFile(ROOT/'tests/fixtures/pandoc-basic.pptx') as z:
        entries = {n: z.read(n) for n in z.namelist() if not n.startswith(('ppt/slides/', 'ppt/notesSlides/', 'ppt/jdecontrols/'))}
    types = ET.fromstring(entries['[Content_Types].xml'])
    for item in list(types):
        name = item.get('PartName', '')
        if name.startswith(('/ppt/slides/', '/ppt/notesSlides/', '/ppt/jdecontrols/')): types.remove(item)
        elif name == '/ppt/presentation.xml': item.set('ContentType', 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml')
    for name, content in [('/ppt/slides/slide1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'),
                          ('/ppt/JDEData.bin', 'application/vnd.wps-officedocument.jdeData')]:
        ET.SubElement(types, f'{{{CT}}}Override', PartName=name, ContentType=content)
    rels = ET.fromstring(entries['ppt/_rels/presentation.xml.rels'])
    for item in list(rels):
        if item.get('Type', '').endswith('/slide'): rels.remove(item)
    ET.SubElement(rels, f'{{{REL}}}Relationship', Id='rIdUserFormGuide', Type=R+'/slide', Target='slides/slide1.xml')
    ET.SubElement(rels, f'{{{REL}}}Relationship', Id='rIdJDE', Type='http://www.wps.cn/officeDocument/2018/jdeExtension', Target='JDEData.bin')
    presentation = ET.fromstring(entries['ppt/presentation.xml'])
    slides = presentation.find(f'{{{P}}}sldIdLst'); slides.clear()
    ET.SubElement(slides, f'{{{P}}}sldId', {'id':'256', f'{{{R}}}id':'rIdUserFormGuide'})
    page_width = int(presentation.find(f'{{{P}}}sldSz').get('cx'))
    text_width = page_width-1200000
    text = shape(2, '表格样式编辑器 · UserForm '+VERSION, 600000, 500000, text_width, 650000, 3000)
    for number, line in enumerate(['启用 JS 宏，在宏编辑器运行 OpenTableStyleEditor。',
            '编辑器与色盘均为独立 UserForm；本页仅提供使用说明。',
            '打开 PPTX 后编辑并保存，手动重新打开目标文稿查看新样式。',
            '当前文稿需先保存正文。恢复备份保存在目标文件旁。'], 3):
        text += shape(number, line, 600000, (number-2)*850000+1000000, text_width, 650000, 1800)
    entries['ppt/slides/slide1.xml'] = f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld name="使用说明"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{text}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'.encode()
    entries['ppt/slides/_rels/slide1.xml.rels'] = f'<Relationships xmlns="{REL}"><Relationship Id="rIdLayout" Type="{R}/slideLayout" Target="../slideLayouts/slideLayout3.xml"/></Relationships>'.encode()
    entries['ppt/JDEData.bin'] = ET.tostring(build_project(bundle), encoding='utf-8', xml_declaration=True)
    for name, node in [('[Content_Types].xml', types), ('ppt/_rels/presentation.xml.rels', rels), ('ppt/presentation.xml', presentation)]:
        entries[name] = ET.tostring(node, encoding='utf-8', xml_declaration=True)
    app = ET.fromstring(entries['docProps/app.xml'])
    for key, value in [('Slides','1'), ('Notes','0')]:
        node = app.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'+key)
        if node is not None: node.text = value
    # These cached lists describe the fixture's original pages, not this guide.
    for key in ['HeadingPairs', 'TitlesOfParts']:
        node = app.find('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'+key)
        if node is not None: app.remove(node)
    entries['docProps/app.xml'] = ET.tostring(app, encoding='utf-8', xml_declaration=True)
    return entries


def build_macro():
    release = ROOT/'release'; release.mkdir(exist_ok=True)
    output = release/f'wps-table-style-macro-{VERSION}.pptm'
    with ZipFile(output, 'w', ZIP_DEFLATED) as z:
        for name, data in build_entries((ROOT/'build/userform-macro-runtime.js').read_text()).items(): z.writestr(name, data)
    archive = release/f'wps-table-style-macro-{VERSION}.zip'
    with ZipFile(archive, 'w', ZIP_DEFLATED) as z:
        z.write(output, '表格样式编辑器-UserForm-'+VERSION+'.pptm')
        z.write(ROOT/'macro/USERFORM-README.md', '使用说明.md')
        for package, license in [('jszip','LICENSE.markdown'), ('@xmldom/xmldom','LICENSE')]:
            z.write(ROOT/'node_modules'/package/license, package.split('/')[-1]+'-LICENSE.txt')
    print(output); print(archive)


if __name__ == '__main__': build_macro()
