"""UserForm resources and its launcher page for the combined macro deck."""
import xml.etree.ElementTree as ET
from form import definitions as original_definitions, control as form_control
from slide import control as native_control


def definitions():
    # Keep the established WPS UserForm property structure; use a fixed layout.
    items = [i for i in original_definitions()
             if i['name'] not in ('ZoomOut', 'ZoomLevel', 'ZoomIn', 'FillChoice', 'TextColorChoice')]
    for item in items:
        if item['kind'] == 'textedit': item['single_line'] = True
        if item['name'] in ('Fill', 'TextColor') or item['name'].startswith('Color_'):
            item['kind'] = 'combobox'
            item['color_field'] = True
            if item['name'] == 'Fill': item['w'] = 188
            elif item['name'] == 'TextColor': item['w'] = 232
            else: item['w'] = 110
        if item['name'].startswith('Width_'):
            item['w'] = 58
            item['action'] = 'border:'+item['name'][6:]
        elif item['name'].startswith(('Color_', 'Dash_')):
            item['action'] = 'border:'+item['name'].split('_', 1)[1]
        elif item['name'].startswith('Mode_'):
            item['action'] = 'mode:'+item['name'][5:]
        if item['name'] == 'DefaultFontSize': item['w'] = 58
    for field in [i for i in items if i.get('color_field')]:
        # A single dropdown plus a plain color swatch, matching the slide UI.
        items.append(dict(kind='label', name='Pick'+field['name'],
                          x=field['x']+field['w']+(12 if field['name']=='TextColor' else 6),
                          y=field['y'], w=26, h=26, caption='\u00a0', color_swatch=True,
                          toolbar=False, event='Click', action='pick:'+field['name']))
    for field in [i for i in items if i['name'].startswith('Width_') or i['name'] == 'DefaultFontSize']:
        for direction, caption, dy in [('Up', '▲', 0), ('Down', '▼', 13)]:
            items.append(dict(kind='commandbutton', name='Step'+direction+'_'+field['name'],
                              x=field['x']+60, y=field['y']+dy, w=22, h=13,
                              caption=caption, toolbar=False, event='Click',
                              font=7,
                              action='step:'+field['name']+(':+1' if direction == 'Up' else ':-1')))
    return items


def build_form():
    return build_resource('TableStyleForm', '表格样式编辑器 · UserForm', 700, 762, definitions())


def build_resource(name, caption, width, height, items, modal=False):
    root = form_control('form', name, 0, 0, width, height, caption)
    root.find("property[@type='facade']/show_modal").text = str(modal).lower()
    for index, item in enumerate(items):
        # Use complete WPS-saved control properties, including combo data and
        # text/value fields, rather than relying on native importer defaults.
        replacement = native_control({**item, 'native': item['name'], 'font': item.get('font', 10)})
        font = replacement.find("property[@type='font']")
        others = ET.Element('property', type='others')
        ET.SubElement(others, 'tab_index').text = str(index)
        ET.SubElement(others, 'tab_stop').text = 'false' if item['kind'] == 'label' else 'true'
        replacement.insert(list(replacement).index(font), others)
        geometry = replacement.find("property[@type='geometry']")
        for key in ('width_policy', 'height_policy'):
            ET.SubElement(geometry, key).text = '1'
        root.append(replacement)
    return root, items


def launcher_definitions():
    return [
        dict(kind='label', name='UserFormIntro', native='TsUserFormIntro', page=4,
             x=220, y=220, w=712, h=36, caption='在独立的 UserForm 窗体中编辑表格样式', font=24, align='center'),
        dict(kind='label', name='UserFormDetails', native='TsUserFormDetails', page=4,
             x=220, y=278, w=712, h=28, caption='打开 PPTX · 新建 / 复制 · 13 个区域 · 六方向边框 · 保存', font=14, align='center'),
        dict(kind='commandbutton', name='OpenUserForm', native='TsOpenUserForm', page=4,
             x=416, y=350, w=320, h=56, caption='打开 UserForm 窗体', font=18, primary=True),
        dict(kind='label', name='UserFormHelp', native='TsUserFormHelp', page=4,
             x=160, y=450, w=832, h=28, caption='启用 JS 宏并退出设计模式，按 Shift+F5 从本页放映，然后点击按钮。', font=12, align='center'),
        dict(kind='label', name='UserFormMacroHelp', native='TsUserFormMacroHelp', page=4,
             x=160, y=490, w=832, h=28, caption='也可在宏编辑器运行 OpenTableStyleUserForm；原幻灯片编辑器保留在第一页。', font=12, align='center'),
    ]


def event_code(items):
    source = '\n'.join(
        f'function TableStyleForm_{i["name"]}_{i["event"]}() {{ return TableStylesUserFormAction("{i["action"]}"); }}'
        for i in items if 'action' in i)
    for item in items:
        if item['name'].startswith('Width_') or item['name'] == 'DefaultFontSize':
            source += f'\nfunction TableStyleForm_{item["name"]}_KeyUp(keycode, shift) {{ var key = Number(keycode); if (key === 38 || key === 16777235) return TableStylesUserFormAction("step:{item["name"]}:+1"); if (key === 40 || key === 16777237) return TableStylesUserFormAction("step:{item["name"]}:-1"); }}'
    return source
