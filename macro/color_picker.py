"""Continuous picker controls on the tool slide (no tiled UserForm)."""

PLANE = dict(x=296, y=150, w=560, h=280)
# Put the two swatches on their own row and the hue strip on the next row.
HUE = dict(x=296, y=500, w=560, h=24)


def definitions(page=2):
    if page not in (2, 3): raise ValueError('Color page must be 2 or 3')
    items = []
    def add(kind, name, x, y, w, h=24, caption='', action=None, event='Click', **extra):
        item = dict(kind=kind, name=name, native=('TsPicker' if page == 2 else 'TsPickerB')+name,
                    x=x, y=y, w=w, h=h, caption=caption, font=12, picker=True, page=page, **extra)
        if action: item.update(action=action, event=event)
        items.append(item)
    add('label', 'Title', 296, 80, 560, 32, '选择颜色', font_override=20)
    add('label', 'Hint', 296, 120, 560, 22, '拖动后松开鼠标取色；输入 HEX / RGB 后按回车预览。', font_override=11)
    add('label', 'Plane', **PLANE, transparent=True, surface='plane')
    add('label', 'Hue', **HUE, transparent=True, surface='hue')
    add('label', 'Original', 296, 450, 48, 24)
    add('label', 'Current', 368, 450, 48, 24)
    add('label', 'OriginalLabel', 296, 480, 48, 20, '原色', font_override=10)
    add('label', 'CurrentLabel', 368, 480, 48, 20, '新色', font_override=10)
    for name, caption, x, w in [('Hex', 'HEX', 296, 164), ('R', 'R', 480, 112), ('G', 'G', 612, 112), ('B', 'B', 744, 112)]:
        add('textedit', name, x, 536, w, 32, action='hex' if name == 'Hex' else 'rgb', event='Change', single_line=True, font_override=10)
        add('label', name+'Label', x, 572, w, 22, caption, align='center', font_override=11)
    add('label', 'ErrorText', 296, 604, 560, 24, foreground='#B42318', font_override=11)
    add('commandbutton', 'Cancel', 632, 646, 104, 32, '取消', 'cancel')
    add('commandbutton', 'Apply', 752, 646, 104, 32, '使用颜色', 'apply', primary=True)
    for item in items:
        item['font'] = item.pop('font_override', 12)
    return items


def event_code(items):
    import json
    code = 'function TableStylesPickerControls(page) {\n    switch (Number(page)) {\n'
    for page in sorted({i['page'] for i in items}):
        controls = ',\n'.join(f'        {i["name"]}: {i["native"]}' for i in items if i['page'] == page)
        code += f'    case {page}: return {{\n'+controls+'\n    };\n'
    code += '    default: throw new Error("缺少色盘页控件：" + page);\n    }\n}\n'
    for i in items:
        if 'surface' in i:
            for event, phase in [('MouseDown', 'down'), ('MouseMove', 'move'), ('MouseUp', 'up')]:
                payload = '{surface: '+json.dumps(i['surface'])+', control: '+json.dumps(i['name'])+', phase: '+json.dumps(phase)+', button: btn, x: x, y: y}'
                code += f'function {i["native"]}_{event}(btn, shift, x, y) {{ return TableStylesColorAction("pointer", {payload}, {i["page"]}); }}\n'
        elif 'action' in i:
            code += f'function {i["native"]}_{i["event"]}() {{ return TableStylesColorAction({json.dumps(i["action"])}, null, {i["page"]}); }}\n'
            if i['kind'] == 'textedit':
                code += f'function {i["native"]}_KeyUp(keycode, shift) {{ if (Number(keycode) === 13 || Number(keycode) === 16777220 || Number(keycode) === 16777221) return TableStylesColorAction("commit", null, {i["page"]}); }}\n'
    return code
