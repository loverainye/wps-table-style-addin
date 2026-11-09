#!/usr/bin/env python3
"""Build a standalone PPTM with native controls embedded on separate editor/picker slides."""
from pathlib import Path
import json
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT/'macro'))
sys.path.insert(0, str(ROOT/'scripts'))
from slide import P, A, R, JDE, EMU, WIDTH, HEIGHT, BACKGROUND, definitions, control_part, picture, event_code, PreviewRenderer
from package_macro import shape, validate_project
from color_picker import definitions as color_definitions, event_code as color_events, PLANE, HUE
from userform import definitions as userform_definitions, build_form as build_userform, launcher_definitions, event_code as userform_events
from userform_color_picker import build_form as build_color_form, event_code as userform_color_events, geometry as userform_color_geometry

VERSION = json.loads((ROOT/'package.json').read_text())['version']
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'


def picker_shapes():
    """Keep the gradients below native hit controls, ready before playback."""
    parts = [('TsColorPlaneImage', PLANE, 'rect', None),
             ('TsColorHueImage', HUE, 'rect', None)]

    for number, (name, geometry, kind, line_color) in enumerate(parts, 3):
        node = ET.Element(f'{{{P}}}sp')
        nv = ET.SubElement(node, f'{{{P}}}nvSpPr')
        ET.SubElement(nv, f'{{{P}}}cNvPr', id=str(number), name=name)
        ET.SubElement(nv, f'{{{P}}}cNvSpPr')
        ET.SubElement(nv, f'{{{P}}}nvPr')
        props = ET.SubElement(node, f'{{{P}}}spPr')
        transform = ET.SubElement(props, f'{{{A}}}xfrm')
        ET.SubElement(transform, f'{{{A}}}off', x=str(geometry['x']*EMU), y=str(geometry['y']*EMU))
        ET.SubElement(transform, f'{{{A}}}ext', cx=str(geometry['w']*EMU), cy=str(geometry['h']*EMU))
        ET.SubElement(ET.SubElement(props, f'{{{A}}}prstGeom', prst=kind), f'{{{A}}}avLst')
        if line_color:
            ET.SubElement(props, f'{{{A}}}noFill')
        else:
            ET.SubElement(ET.SubElement(props, f'{{{A}}}solidFill'), f'{{{A}}}srgbClr', val='FFFFFF')
        line = ET.SubElement(props, f'{{{A}}}ln', w=str(round(1.5*EMU)))
        if line_color:
            ET.SubElement(ET.SubElement(line, f'{{{A}}}solidFill'), f'{{{A}}}srgbClr', val=line_color)
        else:
            ET.SubElement(line, f'{{{A}}}noFill')
        yield node


def userform_code():
    return (ROOT/'macro/userform-code.js').read_text().replace(
        '__TABLE_STYLES_USERFORM_GEOMETRY__', json.dumps(userform_definitions(), ensure_ascii=False)
    ).replace('__TABLE_STYLES_USERFORM_PICKER_GEOMETRY__', json.dumps(userform_color_geometry()))+'\n'+userform_events(userform_definitions())+'\n'+userform_color_events()


def runtime_code(bundle):
    # Load the ZIP/XML library once, on demand, for both interfaces.
    return ('function TableStylesNativeRuntime() {\n'
            'if (!TableStylesNativeRuntime.cached) {\n'+bundle+
            '\nTableStylesNativeRuntime.cached = NativeSlideTableStyles;\n}\n'
            'return TableStylesNativeRuntime.cached;\n}\n')


def build_project(source, runtime='', events=''):
    project = ET.Element('document', version='2.0')
    ET.SubElement(project, 'name').text = 'TableStylesSlide'
    ET.SubElement(project, 'property', desc='', lock='false', password='')
    ET.SubElement(project, 'activemodule').text = '1'
    module = ET.SubElement(project, 'codemodule', name='TableStyleEditor', id='1')
    cursor = source.index('function OpenTableStyleEditor() {')+len('function OpenTableStyleEditor() {')
    ET.SubElement(module, 'window', cursorpos=str(cursor), actived='true', visible='false')
    ET.SubElement(module, 'codetext').text = source
    for name, number, text in [('TableStyleEvents', '2', events), ('TableStyleRuntime', '3', runtime),
                               ('TableStyleUserForm', '4', userform_code())]:
        extra = ET.SubElement(project, 'codemodule', name=name, id=number)
        ET.SubElement(extra, 'window', cursorpos='0', actived='false', visible='false')
        ET.SubElement(extra, 'codetext').text = text
    dialog = ET.SubElement(project, 'dlgmodule', name='TableStyleForm', id='5')
    ET.SubElement(dialog, 'window', visible='false')
    dialog.append(build_userform()[0])
    picker_dialog = ET.SubElement(project, 'dlgmodule', name='TableStyleColorForm', id='6')
    ET.SubElement(picker_dialog, 'window', visible='false')
    picker_dialog.append(build_color_form()[0])
    ET.SubElement(project, 'functionsdata')
    validate_project(project)
    return project


def build_entries(source, render, runtime='', events=''):
    with ZipFile(ROOT/'tests/fixtures/pandoc-basic.pptx') as z:
        entries = {name: z.read(name) for name in z.namelist()}
    types = ET.fromstring(entries['[Content_Types].xml'])
    for item in list(types):
        if re.fullmatch('/ppt/slides/slide[0-9]+.xml', item.get('PartName', '')):
            types.remove(item)
        if item.get('PartName') == '/ppt/presentation.xml':
            item.set('ContentType', 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml')
    for name in list(entries):
        if re.fullmatch('ppt/slides/(?:_rels/)?slide[0-9]+.xml(?:.rels)?', name): del entries[name]
    ET.SubElement(types, f'{{{CT}}}Override', PartName='/ppt/JDEData.bin', ContentType='application/vnd.wps-officedocument.jdeData')
    if not any(x.get('Extension') == 'png' for x in types):
        ET.SubElement(types, f'{{{CT}}}Default', Extension='png', ContentType='image/png')
    presentation_rels = ET.fromstring(entries['ppt/_rels/presentation.xml.rels'])
    for item in list(presentation_rels):
        if item.get('Type', '').endswith('/slide'): presentation_rels.remove(item)
    ET.SubElement(presentation_rels, f'{{{REL}}}Relationship', Id='rIdJDE', Type='http://www.wps.cn/officeDocument/2018/jdeExtension', Target='JDEData.bin')
    presentation = ET.fromstring(entries['ppt/presentation.xml'])
    ids = presentation.find(f'{{{P}}}sldIdLst')
    for item in list(ids): ids.remove(item)
    size = presentation.find(f'{{{P}}}sldSz')
    size.attrib = dict(cx=str(WIDTH*EMU), cy=str(HEIGHT*EMU), type='custom')
    number = 0
    for page, items in [(1, definitions()), (2, color_definitions(2)), (3, color_definitions(3)),
                        (4, launcher_definitions())]:
        slide_rel = f'rIdStyleEditorSlide{page}'
        ET.SubElement(presentation_rels, f'{{{REL}}}Relationship', Id=slide_rel, Type=R+'/slide', Target=f'slides/slide{page}.xml')
        ET.SubElement(ids, f'{{{P}}}sldId', {'id': str(255+page), f'{{{R}}}id': slide_rel})
        ET.SubElement(types, f'{{{CT}}}Override', PartName=f'/ppt/slides/slide{page}.xml', ContentType='application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
        rels = ET.Element(f'{{{REL}}}Relationships')
        ET.SubElement(rels, f'{{{REL}}}Relationship', Id='rIdLayout', Type=R+'/slideLayout', Target='../slideLayouts/slideLayout3.xml')
        slide = ET.Element(f'{{{P}}}sld')
        page_title = 'UserForm 用户窗体版' if page == 4 else '表格样式编辑器'
        cs = ET.SubElement(slide, f'{{{P}}}cSld', name=page_title if page in (1, 4) else f'选择颜色 {page-1}')
        bg = ET.SubElement(ET.SubElement(cs, f'{{{P}}}bg'), f'{{{P}}}bgPr')
        ET.SubElement(ET.SubElement(bg, f'{{{A}}}solidFill'), f'{{{A}}}srgbClr', val=BACKGROUND[1:])
        ET.SubElement(bg, f'{{{A}}}effectLst')
        tree = ET.SubElement(cs, f'{{{P}}}spTree')
        nv = ET.SubElement(tree, f'{{{P}}}nvGrpSpPr')
        ET.SubElement(nv, f'{{{P}}}cNvPr', id='1', name='')
        ET.SubElement(nv, f'{{{P}}}cNvGrpSpPr'); ET.SubElement(nv, f'{{{P}}}nvPr')
        ET.SubElement(tree, f'{{{P}}}grpSpPr')
        title = shape(2, page_title, 24*EMU, 16*EMU, 560*EMU, 38*EMU, 2200)
        title_node = ET.fromstring(f'<root xmlns:p="{P}" xmlns:a="{A}">{title}</root>')[0]
        body = title_node.find(f'{{{P}}}txBody/{{{A}}}bodyPr')
        body.attrib.update(lIns='0', rIns='0', tIns='0', bIns='0')
        style = title_node.find(f'.//{{{A}}}rPr')
        style.set('b', '1')
        style.find(f'{{{A}}}solidFill/{{{A}}}srgbClr').set('val', '20242A')
        for face in ['latin', 'ea']: ET.SubElement(style, f'{{{A}}}{face}', typeface='Noto Sans CJK SC')
        ET.SubElement(ET.SubElement(title_node.find(f'{{{P}}}spPr'), f'{{{A}}}ln'), f'{{{A}}}noFill')
        tree.append(title_node)
        if page in (2, 3): tree.extend(picker_shapes())
        for local_number, item in enumerate(items, 9):
            number += 1
            control_id, image_id = f'rIdControl{number}', f'rIdImage{number}'
            control_name, image_name = f'jdecontrol{number}.xml', f'style-control{number}.png'
            entries['ppt/jdecontrols/'+control_name] = control_part(item)
            entries['ppt/media/'+image_name] = render(item)
            ET.SubElement(types, f'{{{CT}}}Override', PartName='/ppt/jdecontrols/'+control_name,
                          ContentType='application/vnd.wps-officedocument.jdeControl+xml')
            ET.SubElement(rels, f'{{{REL}}}Relationship', Id=control_id, Type=JDE, Target='../jdecontrols/'+control_name)
            ET.SubElement(rels, f'{{{REL}}}Relationship', Id=image_id, Type=R+'/image', Target='../media/'+image_name)
            tree.append(picture(item, local_number, image_id, control_id))
        ET.SubElement(ET.SubElement(slide, f'{{{P}}}clrMapOvr'), f'{{{A}}}masterClrMapping')
        ET.SubElement(slide, f'{{{P}}}transition', advClick='0')
        entries[f'ppt/slides/slide{page}.xml'] = ET.tostring(slide, encoding='utf-8', xml_declaration=True)
        entries[f'ppt/slides/_rels/slide{page}.xml.rels'] = ET.tostring(rels, encoding='utf-8', xml_declaration=True)
    entries['ppt/_rels/presentation.xml.rels'] = ET.tostring(presentation_rels, encoding='utf-8', xml_declaration=True)
    entries['ppt/presentation.xml'] = ET.tostring(presentation, encoding='utf-8', xml_declaration=True)
    entries['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8', xml_declaration=True)
    entries['ppt/JDEData.bin'] = ET.tostring(build_project(source, runtime, events), encoding='utf-8', xml_declaration=True)
    validate_project(ET.fromstring(entries['ppt/JDEData.bin']))
    return entries



def picker_preview(renderer):
    """Design preview, using the same gradient geometry as the macro renderer."""
    import colorsys
    from color_picker import PLANE, HUE
    items = definitions(True)
    im = renderer.page(items, picker=True)
    hsv = colorsys.rgb_to_hsv(116/255, 196/255, 69/255)
    def gradient(geometry, strip=False):
        w, h = geometry['w']*2, geometry['h']*2
        tile = renderer.Image.new('RGB', (w, h))
        pixels = tile.load()
        for y in range(h):
            for x in range(w):
                rgb = colorsys.hsv_to_rgb(1-x/(w-1), 1, 1) if strip else colorsys.hsv_to_rgb(hsv[0], x/(w-1), 1-y/(h-1))
                pixels[x,y] = tuple(round(c*255) for c in rgb)
        im.paste(tile, (geometry['x']*2, geometry['y']*2))
    gradient(PLANE); gradient(HUE, True)
    draw = renderer.ImageDraw.Draw(im)
    for name, text in [('Hex', '#74C445'), ('R', '116'), ('G', '196'), ('B', '69')]:
        item = next(i for i in items if i.get('picker') and i['name']==name)
        tile = renderer.Image.open(__import__('io').BytesIO(renderer.render({**item, 'hidden': False, 'text': text})))
        im.paste(tile, (item['x']*2, item['y']*2))
    for name, color in [('Original', '#C00000'), ('Current', '#74C445')]:
        item = next(i for i in items if i.get('picker') and i['name']==name)
        draw.rectangle((item['x']*2, item['y']*2, (item['x']+item['w'])*2-1, (item['y']+item['h'])*2-1), fill=color)
    return im


def userform_picker_preview(renderer):
    """Design preview of the actual native controls, without PPT shapes."""
    from userform_color_picker import definitions, WIDTH, HEIGHT, COLS, ROWS, hex_color
    import colorsys
    import io
    top = 28
    im = renderer.Image.new('RGB', (WIDTH*2, (HEIGHT+top)*2), BACKGROUND)
    draw = renderer.ImageDraw.Draw(im)
    draw.text((24, 12), '选择颜色 · UserForm', font=renderer.font(24), fill='#20242A')
    hue = colorsys.rgb_to_hsv(116/255, 196/255, 69/255)[0]
    values = {'Hex': '#74C445', 'R': '116', 'G': '196', 'B': '69'}
    for original in definitions():
        item = dict(original)
        if item.get('paint_cell') == 'plane':
            row, col = map(int, item['name'].split('_')[1:])
            item['background'] = hex_color(hue, col/(COLS-1), 1-row/(ROWS-1))
        if item['name'] in values: item['text'] = values[item['name']]
        if item['name'] == 'Current': item['background'] = '#74C445'
        tile = renderer.Image.open(io.BytesIO(renderer.render(item)))
        im.paste(tile, (item['x']*2, (item['y']+top)*2), tile if tile.mode == 'RGBA' else None)
    return im


def build_macro():
    release = ROOT/'release/slide-macro'
    release.mkdir(parents=True, exist_ok=True)
    bundle = (ROOT/'build/slide-macro-runtime.js').read_text()
    source = (ROOT/'macro/slide-code.js').read_text().replace('__TABLE_STYLES_VERSION__', json.dumps(VERSION))
    runtime = runtime_code(bundle)
    events = event_code(definitions())+'\n'+color_events(color_definitions(2)+color_definitions(3))
    renderer = PreviewRenderer()
    entries = build_entries(source, renderer.render, runtime, events)
    output = release/f'表格样式编辑器-幻灯片宏版-{VERSION}.pptm'
    with ZipFile(output, 'w', ZIP_DEFLATED) as z:
        for name, data in entries.items(): z.writestr(name, data)
    # GitHub normalizes non-ASCII asset names. Keep the Chinese name in the ZIP
    # and provide a stable ASCII filename for the standalone release download.
    download = ROOT/'release'/f'wps-table-style-macro-{VERSION}.pptm'
    download.write_bytes(output.read_bytes())
    (release/'编辑器入口.js').write_text(source)
    (release/'控件事件.js').write_text(events)
    (release/'离线运行库.js').write_text(runtime)
    (release/'UserForm入口.js').write_text(userform_code())
    (release/'UserForm窗体.xml').write_bytes(ET.tostring(build_userform()[0], encoding='utf-8', xml_declaration=True))
    (release/'UserForm色盘.xml').write_bytes(ET.tostring(build_color_form()[0], encoding='utf-8', xml_declaration=True))
    (release/'使用说明.md').write_text((ROOT/'macro/SLIDE-README.md').read_text())
    renderer.page(definitions()).save(release/'界面预览.png')
    picker_preview(renderer).save(release/'色盘预览.png')
    renderer.page(launcher_definitions(), title='UserForm 用户窗体版').save(release/'UserForm启动页预览.png')
    userform_picker_preview(renderer).save(release/'UserForm色盘预览.png')
    for package, filename in [('jszip', 'LICENSE.markdown'), ('@xmldom/xmldom', 'LICENSE')]:
        (release/(package.split('/')[-1]+'-LICENSE.txt')).write_bytes((ROOT/'node_modules'/package/filename).read_bytes())
    archive = ROOT/'release'/f'wps-table-style-slide-macro-{VERSION}.zip'
    with ZipFile(archive, 'w', ZIP_DEFLATED) as z:
        files = [output, *(release/name for name in ['编辑器入口.js', '控件事件.js', '离线运行库.js',
                                                   'UserForm入口.js', 'UserForm窗体.xml', 'UserForm色盘.xml', 'UserForm启动页预览.png', 'UserForm色盘预览.png',
                                                   '使用说明.md', '界面预览.png', '色盘预览.png', 'jszip-LICENSE.txt', 'xmldom-LICENSE.txt'])]
        for path in files: z.write(path, path.name)
    print(output); print(download); print(archive)


if __name__ == '__main__': build_macro()
