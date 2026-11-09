"""Native slide controls, matching the user's WPS-saved jdecontrol sample."""
from copy import deepcopy
from pathlib import Path
import io
import subprocess
import xml.etree.ElementTree as ET

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
JDE = 'http://www.wps.cn/officeDocument/2020/jdeControlExtension'
EMU = 12700
WIDTH, HEIGHT = 1152, 720
BACKGROUND = '#F5F6F8'
TEMPLATES = Path(__file__).with_name('slide-control-templates.xml')
for prefix, uri in [('p', P), ('a', A), ('r', R), ('wpsjdecontrol', JDE)]:
    ET.register_namespace(prefix, uri)


def definitions(include_picker=False):
    items = []
    def add(kind, name, x, y, w, h=28, caption='', action=None, event='Click', **extra):
        item = dict(kind=kind, name=name, native='Ts'+name, page=1, x=x, y=y, w=w, h=h,
                    caption=caption, font=12)
        item.update(extra)
        if action: item.update(action=action, event=event)
        items.append(item)
    add('commandbutton', 'OpenFile', 850, 18, 122, 32, '打开 PPTX', 'open')
    add('commandbutton', 'SaveStyle', 984, 18, 144, 32, '保存并应用', 'save', primary=True)
    add('label', 'FirstPlaceholder', 24, 60, 1104, 24,
        '启用宏并退出设计模式后，按 F5 开始放映，再点击“打开 PPTX”。', font=11)
    add('label', 'StyleLabel', 24, 94, 368, 22, '表格样式')
    add('combobox', 'Styles', 24, 120, 368, action='style', event='Change')
    add('combobox', 'Preset', 24, 160, 112, text='绿色')
    add('commandbutton', 'NewStyle', 148, 160, 110, caption='新建', action='new')
    add('commandbutton', 'DuplicateStyle', 270, 160, 122, caption='复制样式', action='duplicate')
    add('label', 'NameLabel', 24, 206, 368, 22, '样式名称')
    add('textedit', 'StyleName', 24, 232, 368, action='fields', event='Change')
    add('label', 'PreviewTitle', 24, 280, 368, 22, '配色预览')
    data = [['项目', '目标', '完成', '达成率'], ['交付质量', '98%', '99%', '101%'],
            ['项目交付', '12项', '12项', '100%'], ['季度合计', '—', '已完成', '100%']]
    for row, values in enumerate(data):
        for col, value in enumerate(values):
            add('label', f'Preview_{row}_{col}', 24+92*col, 308+32*row,
                91, 31, value, align='center', background='#E5E9EE' if row == 0 else '#FFFFFF', font=11)
    add('label', 'OptionsTitle', 24, 456, 368, 22, '表格选项')
    flags = ['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol']
    for i, (key, title) in enumerate(zip(flags, ['标题行', '汇总行', '第一列', '最后一列', '条纹行', '条纹列'])):
        add('checkbox', 'Flag_'+key, 24+(i % 3)*124, 482+(i // 3)*30, 120,
            caption=title, action='flags')
    add('label', 'ScopeLabel', 24, 552, 368, 22, '应用范围')
    add('combobox', 'Scope', 24, 578, 368, text='所有表格')
    add('checkbox', 'SetDefault', 24, 616, 368, caption='设为文稿默认表格样式')
    add('label', 'PartLabel', 420, 94, 708, 22, '编辑区域')
    add('combobox', 'Part', 420, 120, 708, text='标题行', action='part', event='Change')
    add('label', 'FillLabel', 420, 160, 338, 22, '填充颜色')
    add('combobox', 'Fill', 420, 186, 290, text='继承', action='fields', event='Change')
    add('label', 'PickFill', 722, 186, 36, caption='\u00a0', action='pick:Fill', color_swatch=True, background='#FFFFFF')
    add('label', 'TextColorLabel', 784, 160, 344, 22, '文字颜色')
    add('combobox', 'TextColor', 784, 186, 296, text='继承', action='fields', event='Change')
    add('label', 'PickTextColor', 1092, 186, 36, caption='\u00a0', action='pick:TextColor', color_swatch=True, background='#FFFFFF')
    add('label', 'BoldLabel', 420, 226, 338, 22, '加粗')
    add('combobox', 'Bold', 420, 252, 338, text='继承', action='fields', event='Change')
    add('label', 'ItalicLabel', 784, 226, 344, 22, '斜体')
    add('combobox', 'Italic', 784, 252, 344, text='继承', action='fields', event='Change')
    add('label', 'BorderLabel', 420, 294, 150, 22, '边框')
    add('commandbutton', 'AllBorders', 1000, 288, 128, caption='上 → 全部', action='allBorders')
    add('label', 'BorderModeLabel', 504, 326, 132, 22, '方式', font=11)
    add('label', 'BorderColorLabel', 648, 326, 190, 22, '颜色 / theme:…', font=11)
    add('label', 'BorderWidthLabel', 850, 326, 86, 22, '线宽（磅）', font=11)
    add('label', 'BorderDashLabel', 948, 326, 180, 22, '线型', font=11)
    for i, (edge, title) in enumerate(zip(['left', 'right', 'top', 'bottom', 'insideH', 'insideV'],
                                         ['左', '右', '上', '下', '内部横线', '内部竖线'])):
        y = 352+i*34
        add('label', 'Label_'+edge, 420, y, 76, caption=title, font=11)
        add('combobox', 'Mode_'+edge, 504, y, 132, text='继承', action='mode:'+edge, event='Change')
        add('combobox', 'Color_'+edge, 648, y, 146, text='#FFFFFF', action='border:'+edge, event='Change')
        add('label', 'PickColor_'+edge, 802, y, 36, caption='\u00a0', action='pick:Color_'+edge, color_swatch=True, background='#FFFFFF')
        add('textedit', 'Width_'+edge, 850, y, 62, 30, text='1', action='border:'+edge, event='Change', font=10, single_line=True)
        for direction, caption, dy in [('Up', '▲', 0), ('Down', '▼', 15)]:
            add('label', 'Step'+direction+'_Width_'+edge, 916, y+dy, 20, 15, caption,
                action='step:Width_'+edge+(':+1' if direction == 'Up' else ':-1'), font=8, align='center', color_swatch=True)
        add('combobox', 'Dash_'+edge, 948, y, 180, text='实线', action='border:'+edge, event='Change')
    add('checkbox', 'ClearOverrides', 420, 578, 708, caption='清除单元格手动设置的颜色、字体和边框')
    add('checkbox', 'SetDefaultFont', 420, 616, 230, caption='设置母版默认字号', action='font')
    add('textedit', 'DefaultFontSize', 658, 616, 56, 30, text='18', font=10, single_line=True)
    for direction, caption, dy in [('Up', '▲', 0), ('Down', '▼', 15)]:
        add('label', 'Step'+direction+'_DefaultFontSize', 716, 616+dy, 20, 15, caption,
            action='step:DefaultFontSize'+(':+1' if direction == 'Up' else ':-1'), font=8, align='center', color_swatch=True)
    add('label', 'FontUnit', 748, 616, 40, caption='磅')
    add('label', 'ErrorText', 24, 664, 1104, 38, foreground='#B42318', font=11)
    if include_picker:
        from color_picker import definitions as picker_definitions
        items.extend(picker_definitions())
    return items


def control(item):
    templates = ET.parse(TEMPLATES).getroot()
    node = deepcopy(templates.find(f"control[@type='{item['kind']}']"))
    def put(group, key, value):
        parent = node.find(f"property[@type='{group}']")
        target = parent.find(key)
        if target is None: target = ET.SubElement(parent, key)
        target.text = str(value)
    put('facade', 'name', item['native'])
    if item.get('hidden'): put('facade', 'visible', 'false')
    if item.get('transparent'): put('facade', 'back_style', 'false')
    if item.get('single_line'): put('actions', 'word_wrap', 'false')
    if item.get('color_swatch'):
        put('facade', 'border_style', 'true')
        put('facade', 'border_color', int('FFD3D8DF', 16))
    for key, source in [('left', 'x'), ('top', 'y'), ('width', 'w'), ('height', 'h')]:
        put('geometry', key, round(item[source]*EMU))
    background = item.get('background', '#FFFFFF' if item['kind'] in ['textedit', 'combobox'] else BACKGROUND)
    foreground = item.get('foreground', '#20242A')
    if item.get('primary'): background, foreground = '#305D46', '#FFFFFF'
    put('facade', 'back_color', int('FF'+background[1:], 16))
    put('facade', 'fore_color', int('FF'+foreground[1:], 16))
    if node.find("property[@type='facade']/caption") is not None:
        put('facade', 'caption', item['caption'] or '\u00a0')
    if item['kind'] in ['textedit', 'combobox']:
        for key in ['text', 'value']: put('facade', key, item.get('text', ''))
    if item.get('align') == 'center': put('actions', 'text_align', 2)
    font = node.find("property[@type='font']/font")
    font.find('family').text = 'Noto Sans CJK SC'
    font.find('pointsize').text = str(item['font'])
    return node


def control_part(item):
    root = ET.Element(f'{{{JDE}}}jdecontrol')
    root.append(control(item))
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def picture(item, number, image_id, control_id):
    pic = ET.Element(f'{{{P}}}pic')
    nv = ET.SubElement(pic, f'{{{P}}}nvPicPr')
    ET.SubElement(nv, f'{{{P}}}cNvPr', id=str(number), name=item['native'])
    ET.SubElement(nv, f'{{{P}}}cNvPicPr')
    ET.SubElement(nv, f'{{{P}}}nvPr')
    fill = ET.SubElement(pic, f'{{{P}}}blipFill')
    ET.SubElement(fill, f'{{{A}}}blip', {f'{{{R}}}embed': image_id})
    ET.SubElement(ET.SubElement(fill, f'{{{A}}}stretch'), f'{{{A}}}fillRect')
    prop = ET.SubElement(pic, f'{{{P}}}spPr')
    xfrm = ET.SubElement(prop, f'{{{A}}}xfrm')
    ET.SubElement(xfrm, f'{{{A}}}off', x=str(round(item['x']*EMU)), y=str(round(item['y']*EMU)))
    ET.SubElement(xfrm, f'{{{A}}}ext', cx=str(round(item['w']*EMU)), cy=str(round(item['h']*EMU)))
    ET.SubElement(ET.SubElement(prop, f'{{{A}}}prstGeom', prst='rect'), f'{{{A}}}avLst')
    # WPS writes jdecontrolRef directly in extLst, without a standard a:ext.
    ext = ET.SubElement(prop, f'{{{A}}}extLst')
    ET.SubElement(ext, f'{{{JDE}}}jdecontrolRef', {f'{{{R}}}id': control_id})
    return pic


def event_code(items):
    items = [i for i in items if not i.get('picker')]
    controls = ',\n'.join(f'        {i["name"]}: {i["native"]}' for i in items)
    source = 'function TableStylesSlideControls(page) {\n'
    source += '    if (Number(page || 1) !== 1) throw new Error("缺少表单页控件：" + page);\n'
    source += '    if (!TableStylesSlideControls.cached) TableStylesSlideControls.cached = {\n'+controls+'\n    };\n'
    source += '    return TableStylesSlideControls.cached;\n}\n'
    opening = ', '.join(f'{i["name"]}: {i["native"]}' for i in items if i['name'] in ('OpenFile', 'FirstPlaceholder'))
    source += 'function TableStylesOpenControls(page) {\n'
    source += '    if (Number(page || 1) !== 1) throw new Error("缺少打开按钮：" + page);\n'
    source += '    return {'+opening+'};\n}\n'
    for item in items:
        name, page = item['native'], item.get('page', 1)
        if item.get('action') == 'open':
            source += f'function {name}_Click() {{ return TableStylesOpenFile({page}); }}\n'
        elif 'action' in item:
            source += f'function {name}_{item["event"]}() {{ return TableStylesSlideAction("{item["action"]}", {page}); }}\n'
        if item['name'] in ['Fill', 'TextColor'] or item['name'].startswith('Color_'):
            source += f'function {name}_Enter() {{ return TableStylesSlideAction("colors:{item["name"]}", {page}); }}\n'
        if item['name'].startswith('Width_') or item['name'] == 'DefaultFontSize':
            source += f'function {name}_KeyUp(keycode, shift) {{ var key = Number(keycode); if (key === 38 || key === 16777235) return TableStylesSlideAction("step:{item["name"]}:+1", {page}); if (key === 40 || key === 16777237) return TableStylesSlideAction("step:{item["name"]}:-1", {page}); }}\n'
    return source


class PreviewRenderer:
    """Render control fallback pictures; WPS replaces these with live controls."""
    def __init__(self):
        from PIL import Image, ImageDraw, ImageFont
        self.Image, self.ImageDraw, self.ImageFont = Image, ImageDraw, ImageFont
        self.font_path = subprocess.check_output(['fc-match', '-f', '%{file}', 'Noto Sans CJK SC'], text=True).strip()
        self.fonts = {}

    def font(self, size, weight='Regular'):
        key = (size, weight)
        if key not in self.fonts:
            font = self.ImageFont.truetype(self.font_path, size)
            try: font.set_variation_by_name(weight)
            except (AttributeError, OSError, ValueError): pass
            self.fonts[key] = font
        return self.fonts[key]

    def render(self, item):
        scale = 2
        w, h = round(item['w']*scale), round(item['h']*scale)
        if item.get('hidden'):
            output = io.BytesIO(); self.Image.new('RGBA', (w, h), (0, 0, 0, 0)).save(output, format='PNG')
            return output.getvalue()
        bg = item.get('background', '#FFFFFF' if item['kind'] in ['textedit', 'combobox'] else BACKGROUND)
        fg = item.get('foreground', '#20242A')
        if item.get('primary'): bg, fg = '#305D46', '#FFFFFF'
        im = self.Image.new('RGBA', (w, h), (0, 0, 0, 0) if item.get('transparent') else bg)
        draw = self.ImageDraw.Draw(im)
        kind = item['kind']
        if kind in ['commandbutton', 'textedit', 'combobox']:
            draw.rounded_rectangle((1, 1, w-2, h-2), radius=5*scale, fill=bg,
                                   outline=bg if item.get('primary') else '#D3D8DF', width=scale)
        size = round(item['font']*scale)
        font = self.font(size)
        text = item.get('text', '') if kind in ['textedit', 'combobox'] else item['caption']
        left, right = 8*scale, w-8*scale
        if kind == 'combobox':
            x, y = w-15*scale, h//2
            draw.line([(x-4*scale, y-2*scale), (x, y+2*scale), (x+4*scale, y-2*scale)], fill='#657182', width=2)
            right -= 20*scale
        if kind == 'checkbox':
            y = h//2-7*scale
            draw.rounded_rectangle((2*scale, y, 16*scale, y+14*scale), radius=2*scale,
                                   fill='#FFFFFF', outline='#A8B0BC', width=scale)
            left = 24*scale
        if item.get('color_swatch'):
            draw.rectangle((0, 0, w-1, h-1), outline='#A8B0BC', width=scale)
        if kind == 'label': left = 0
        while text and draw.textlength(text, font=font) > right-left:
            text = text[:-2]+'…' if len(text) > 1 else ''
        if kind == 'commandbutton' or item.get('align') == 'center':
            left = (w-draw.textlength(text, font=font))/2
        draw.text((left, h/2), text, font=font, fill=fg, anchor='lm')
        output = io.BytesIO(); im.save(output, format='PNG')
        return output.getvalue()

    def page(self, items, picker=False, title='表格样式编辑器'):
        im = self.Image.new('RGB', (WIDTH*2, HEIGHT*2), BACKGROUND)
        draw = self.ImageDraw.Draw(im)
        font = self.font(44, 'Bold')
        draw.text((48, 36), title, font=font, fill='#20242A')
        for item in items:
            if bool(item.get('picker')) != picker: continue
            image = self.Image.open(io.BytesIO(self.render({**item, 'hidden': False})))
            im.paste(image, (round(item['x']*2), round(item['y']*2)), image if image.mode == 'RGBA' else None)
        return im
