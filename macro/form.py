"""Build WPS JSA native UserForm resources (not HTML or a registered add-in)."""
import xml.etree.ElementTree as ET

EDGES = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']
EMU_PER_POINT = 12700


def definitions():
    items = []
    def add(kind, name, x, y, w, h=26, caption='', action=None, event='Click', toolbar=False):
        item = dict(kind=kind, name=name, x=x, y=y, w=w, h=h, caption=caption, toolbar=toolbar)
        if action: item.update(action=action, event=event)
        items.append(item)
    add('commandbutton', 'OpenFile', 14, 10, 100, caption='打开 PPTX', action='open', toolbar=True)
    add('commandbutton', 'UseCurrent', 122, 10, 100, caption='当前文稿', action='current', toolbar=True)
    add('commandbutton', 'ZoomOut', 394, 10, 32, caption='−', action='zoomOut', toolbar=True)
    add('label', 'ZoomLevel', 432, 12, 52, caption='100%', toolbar=True)
    add('commandbutton', 'ZoomIn', 490, 10, 32, caption='+', action='zoomIn', toolbar=True)
    add('commandbutton', 'SaveStyle', 532, 10, 94, caption='保存并应用', action='save', toolbar=True)
    add('commandbutton', 'CloseEditor', 634, 10, 52, caption='关闭', action='close', toolbar=True)
    add('label', 'FirstPlaceholder', 14, 48, 672, 22)
    add('label', 'StyleLabel', 14, 76, 70, caption='表格样式')
    add('combobox', 'Styles', 92, 72, 308, action='style', event='Change')
    add('combobox', 'Preset', 408, 72, 80)
    add('commandbutton', 'NewStyle', 496, 72, 86, caption='新建', action='new')
    add('commandbutton', 'DuplicateStyle', 590, 72, 96, caption='复制样式', action='duplicate')
    add('label', 'NameLabel', 14, 112, 70, caption='样式名称')
    add('textedit', 'StyleName', 92, 108, 594, action='fields', event='Change')
    data = [['项目', '目标', '完成', '达成率'], ['交付质量', '98%', '99%', '101%'], ['项目交付', '12项', '12项', '100%'], ['季度合计', '—', '已完成', '100%']]
    for r, row in enumerate(data):
        for c, text in enumerate(row):
            add('label', f'Preview_{r}_{c}', 14 + 168*c, 146 + 25*r, 167, 24, text)
    add('label', 'PartLabel', 14, 264, 70, caption='编辑区域')
    add('combobox', 'Part', 92, 260, 594, action='part', event='Change')
    add('label', 'FillLabel', 14, 300, 70, caption='填充颜色')
    add('textedit', 'Fill', 92, 296, 112, action='fields', event='Change')
    add('combobox', 'FillChoice', 210, 296, 102, action='fillChoice', event='Change')
    add('label', 'TextColorLabel', 334, 300, 74, caption='文字颜色')
    add('textedit', 'TextColor', 416, 296, 144, action='fields', event='Change')
    add('combobox', 'TextColorChoice', 566, 296, 120, action='textColorChoice', event='Change')
    add('label', 'BoldLabel', 14, 336, 70, caption='加粗')
    add('combobox', 'Bold', 92, 332, 220, action='fields', event='Change')
    add('label', 'ItalicLabel', 334, 336, 74, caption='斜体')
    add('combobox', 'Italic', 416, 332, 270, action='fields', event='Change')
    add('label', 'BorderLabel', 14, 372, 70, caption='边框')
    add('label', 'BorderModeLabel', 92, 372, 120, caption='方式')
    add('label', 'BorderColorLabel', 224, 372, 144, caption='颜色 / theme:…')
    add('label', 'BorderWidthLabel', 380, 372, 82, caption='线宽（磅）')
    add('label', 'BorderDashLabel', 476, 372, 114, caption='线型')
    add('commandbutton', 'AllBorders', 594, 368, 92, caption='上 → 全部', action='allBorders')
    for i, (edge, title) in enumerate(zip(EDGES, ['左', '右', '上', '下', '内部横线', '内部竖线'])):
        y = 402 + i*30
        add('label', 'Label_' + edge, 14, y + 4, 70, caption=title)
        add('combobox', 'Mode_' + edge, 92, y, 120, action='fields', event='Change')
        add('textedit', 'Color_' + edge, 224, y, 144, action='fields', event='Change')
        add('textedit', 'Width_' + edge, 380, y, 82, action='fields', event='Change')
        add('combobox', 'Dash_' + edge, 476, y, 210, action='fields', event='Change')
    for i, (flag, title) in enumerate(zip(['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol'], ['标题行', '汇总行', '第一列', '最后一列', '条纹行', '条纹列'])):
        add('checkbox', 'Flag_' + flag, 14 + 112*i, 590, 106, caption=title, action='flags')
    add('label', 'ScopeLabel', 14, 630, 70, caption='应用范围')
    add('combobox', 'Scope', 92, 626, 300)
    add('checkbox', 'SetDefault', 408, 626, 278, caption='设为文稿默认表格样式')
    add('checkbox', 'ClearOverrides', 14, 664, 354, caption='清除单元格手动设置的颜色、字体和边框')
    add('checkbox', 'SetDefaultFont', 380, 664, 170, caption='设置母版默认字号', action='font')
    add('textedit', 'DefaultFontSize', 558, 664, 82)
    add('label', 'FontUnit', 648, 668, 38, caption='磅')
    add('label', 'ErrorText', 14, 704, 672, 48)
    return items


def control(kind, name, x, y, w, h, caption='', tabindex=0):
    node = ET.Element('control', type=kind)
    def prop(group, values):
        p = ET.SubElement(node, 'property', type=group)
        for k, v in values.items(): ET.SubElement(p, k).text = str(v)
    if kind != 'form':
        prop('actions', {'enabled': 'true', 'word_wrap': 'true', 'auto_size': 'false'})
    values = {'name': name, 'back_color': 0xFFF3F4F6, 'fore_color': 0xFF20242A}
    if kind in ['form', 'commandbutton', 'label', 'checkbox']: values['caption'] = caption
    if kind in ['textedit', 'combobox']: values.update(text='', back_color=0xFFFFFFFF)
    if name == 'ErrorText': values['fore_color'] = 0xFFB42318
    if kind == 'form':
        values.update(border_style='false', border_color=0xFF20242A,
                      show_modal='false', startup_position=1)
    else:
        values.update(visible='true', back_style='true')
    prop('facade', values)
    # Native JDE XML stores EMU, while the macro's Left/Top/Width/Height API
    # uses points. Writing 700 directly produces a form less than 0.06 pt wide.
    # Verified against output/原生窗体样本.pptm saved by the user's WPS.
    prop('geometry', {**{k: round(v * EMU_PER_POINT) for k, v in
                        dict(left=x, top=y, width=w, height=h).items()},
                      'width_policy': 1, 'height_policy': 1})
    if kind == 'form':
        prop('layout', {'layout_type': 1})
    else:
        prop('others', {'tab_index': tabindex, 'tab_stop': 'false' if kind == 'label' else 'true'})
    font = ET.SubElement(ET.SubElement(node, 'property', type='font'), 'font')
    for k, v in dict(family='微软雅黑', pointsize=10, weight=50, italic='false', underline='false', strickout='false').items():
        ET.SubElement(font, k).text = str(v)
    return node


def build_form(items=None):
    root = control('form', 'TableStyleForm', 0, 0, 700, 762, '表格样式编辑器')
    if items is None: items = definitions()
    for i, item in enumerate(items):
        root.append(control(item['kind'], item['name'], item['x'], item['y'], item['w'], item['h'], item['caption'], i))
    return root, items


def event_code(items):
    return '\n'.join(f'function TableStyleForm_{i["name"]}_{i["event"]}() {{ TableStylesFormEditor().dispatch("{i["action"]}"); }}' for i in items if 'action' in i)
