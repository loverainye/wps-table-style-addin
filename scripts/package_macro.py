#!/usr/bin/env python3
"""Embed a native UserForm and offline runtime in a standalone WPS JSA PPTM."""
from pathlib import Path
import json
import sys
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / 'macro'))
from form import build_form, event_code
VERSION = json.loads((ROOT / 'package.json').read_text())['version']
RELEASE = ROOT / 'release' / 'macro'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def shape(number, text, x, y, w, h, size=2200, macro=None):
    action = f'<a:hlinkClick action="ppaction://macro?name={macro}"/>' if macro else ''
    fill = '<a:solidFill><a:srgbClr val="E8EFE9"/></a:solidFill>' if macro else '<a:noFill/>'
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{number}" name="Text {number}">{action}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}</p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:pPr/><a:r><a:rPr lang="zh-CN" sz="{size}"><a:solidFill><a:srgbClr val="305D46"/></a:solidFill></a:rPr><a:t>{escape(text)}</a:t></a:r><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>'''


def validate_project(project):
    """Reject nodes which the native dialog loader treats as controls.

    Unlike codemodule, dlgmodule does not accept a codetext child. Linux WPS
    sends every non-window child to its control XML importer; passing a text
    node there crashes libjdecontrol before any macro entry point runs.
    """
    for dialog in project.findall('dlgmodule'):
        if [child.tag for child in dialog] != ['window', 'control']:
            raise ValueError('dlgmodule 只能包含 window 和 control；宏代码必须放在 codemodule 中')
        form = dialog.find('control')
        if form.get('type') != 'form':
            raise ValueError('dlgmodule 必须包含一个 form 控件')
        for control in form.iter('control'):
            if any(child.tag not in ('property', 'control') for child in control):
                raise ValueError('原生控件资源中不能嵌入宏代码')
    for text in project.iter('codetext'):
        if not any(text in list(module) for module in project.findall('codemodule')):
            raise ValueError('codetext 必须是标准宏模块的直接子节点')


def build_project(code, form_code, form):
    project = ET.Element('document', version='2.0')
    ET.SubElement(project, 'name').text = 'TableStyles'
    ET.SubElement(project, 'property', desc='', lock='false', password='')
    ET.SubElement(project, 'activemodule').text = '1'
    # Form events are public JSA functions, kept with the launcher in the
    # standard module. The dialog module contains only its native resources.
    module = ET.SubElement(project, 'codemodule', name='TableStyleEditor', id='1')
    source = code + '\n' + form_code
    entry = 'function OpenTableStyleEditor() {'
    cursor = source.find(entry)
    # Position F5 within the actual launcher body, not a return-only probe.
    cursor = cursor + len(entry) if cursor >= 0 else 0
    ET.SubElement(module, 'window', cursorpos=str(cursor), actived='true', visible='false')
    ET.SubElement(module, 'codetext').text = source
    dialog = ET.SubElement(project, 'dlgmodule', name='TableStyleForm', id='2')
    ET.SubElement(dialog, 'window', visible='false')
    dialog.append(form)
    ET.SubElement(project, 'functionsdata')
    validate_project(project)
    return project


def build_macro():
    RELEASE.mkdir(parents=True, exist_ok=True)
    form, controls = build_form()
    bundle = (ROOT / 'build' / 'macro-runtime.js').read_text()
    code = (ROOT / 'macro' / 'launcher.js').read_text().replace('__TABLE_STYLES_VERSION__', json.dumps(VERSION))
    form_code = (ROOT / 'macro' / 'form-code.js').read_text().replace('__TABLE_STYLES_NATIVE_BUNDLE__', bundle).replace('__TABLE_STYLES_NATIVE_GEOMETRY__', json.dumps(controls, ensure_ascii=False)) + '\n' + event_code(controls)
    (RELEASE / '表格样式编辑器.js').write_text(code + '\n' + form_code)
    (RELEASE / '原生窗体代码.js').write_text(form_code)
    (RELEASE / '原生窗体.xml').write_bytes(ET.tostring(form, encoding='utf-8', xml_declaration=True))
    project = build_project(code, form_code, form)
    # Use an independent Pandoc fixture, never the user's reference/master deck.
    with ZipFile(ROOT / 'tests' / 'fixtures' / 'pandoc-basic.pptx') as z:
        entries = {name: z.read(name) for name in z.namelist()}
    types = ET.fromstring(entries['[Content_Types].xml'])
    for item in types:
        if item.get('PartName') == '/ppt/presentation.xml':
            item.set('ContentType', 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml')
    ET.SubElement(types, f'{{{CT}}}Override', PartName='/ppt/JDEData.bin', ContentType='application/vnd.wps-officedocument.jdeData')
    entries['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8', xml_declaration=True)
    relationships = ET.fromstring(entries['ppt/_rels/presentation.xml.rels'])
    ET.SubElement(relationships, f'{{{REL}}}Relationship', Id='rIdJDE', Type='http://www.wps.cn/officeDocument/2018/jdeExtension', Target='JDEData.bin')
    entries['ppt/_rels/presentation.xml.rels'] = ET.tostring(relationships, encoding='utf-8', xml_declaration=True)
    presentation = ET.fromstring(entries['ppt/presentation.xml'])
    slides = presentation.find(f'{{{P}}}sldIdLst')
    for item in list(slides)[1:]: slides.remove(item)
    entries['ppt/presentation.xml'] = ET.tostring(presentation, encoding='utf-8', xml_declaration=True)
    title = shape(2, '表格样式编辑器 · JS 宏版', 600000, 400000, 10000000, 650000, 3200)
    text = shape(3, '启用此文件的 JS 宏，在宏编辑器运行 OpenTableStyleEditor。', 600000, 1400000, 10000000, 600000)
    text += shape(4, '点击原生窗体中的“打开 PPTX”加载文件；无需安装加载项或网页窗体接口。', 600000, 2100000, 10000000, 600000, 2000)
    text += shape(5, '保存只修改文件，不自动关闭或重开。需查看新样式时，手动重新打开目标文稿。', 600000, 2800000, 10000000, 800000, 2000)
    button = shape(6, '打开编辑器（放映模式下点击）', 600000, 4100000, 6500000, 700000, 2400, 'OpenTableStyleEditor')
    entries['ppt/slides/slide1.xml'] = f'''<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld name="JS 宏表格样式编辑器"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><!---->{title}{text}{button}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'''.encode()
    entries['ppt/JDEData.bin'] = ET.tostring(project, encoding='utf-8', xml_declaration=True)
    # Check the embedded source too: XML serializers do not reject literal C0
    # characters in JavaScript strings, but the native JDE XML reader does.
    validate_project(ET.fromstring(entries['ppt/JDEData.bin']))
    output = RELEASE / '表格样式编辑器-宏版.pptm'
    with ZipFile(output, 'w', ZIP_DEFLATED) as z:
        for name, data in entries.items(): z.writestr(name, data)
    (RELEASE / '使用说明.md').write_text((ROOT / 'macro' / 'README.md').read_text())
    (RELEASE / 'JSZip-LICENSE.txt').write_text((ROOT / 'node_modules' / 'jszip' / 'LICENSE.markdown').read_text())
    (RELEASE / 'xmldom-LICENSE.txt').write_text((ROOT / 'node_modules' / '@xmldom' / 'xmldom' / 'LICENSE').read_text())
    archive = ROOT / 'release' / f'wps-table-style-macro-{VERSION}.zip'
    with ZipFile(archive, 'w', ZIP_DEFLATED) as z:
        for path in sorted(RELEASE.iterdir()):
            if path.is_file(): z.write(path, path.name)
    print(output)
    print(archive)


if __name__ == '__main__':
    build_macro()
