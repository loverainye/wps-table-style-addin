"""A self-contained native UserForm color picker; no slide or image API."""
import colorsys
import json
from userform import build_resource

WIDTH, HEIGHT = 420, 526
PLANE = dict(x=18, y=44, w=384, h=216)
HUE = dict(x=18, y=344, w=384, h=20)
COLS, ROWS, HUE_COLS = 32, 24, 64
SWATCHES = ['#FFFFFF', '#000000', '#E7E6E6', '#7F7F7F', '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050',
            '#00B0F0', '#0070C0', '#002060', '#7030A0', '#385723', '#A9D18E', '#F4B183', '#F8CBAD', '#B4C6E7', '#D9E1F2']


def hex_color(h, s, v):
    return '#'+''.join(f'{round(c*255):02X}' for c in colorsys.hsv_to_rgb(h, s, v))


def geometry():
    return dict(plane=PLANE, hue=HUE, columns=COLS, rows=ROWS, hueColumns=HUE_COLS)


def definitions():
    items = []
    def add(kind, name, x, y, w, h, caption='\u00a0', **extra):
        items.append(dict(kind=kind, name=name, x=x, y=y, w=w, h=h, caption=caption, font=10, **extra))
    add('label', 'Hint', 18, 14, 384, 22, '拖动取色；输入 HEX / RGB 后按回车预览。')
    # Plain native labels render the gradients. Only the two transparent
    # overlays receive pointer events, so event code stays small.
    for row in range(ROWS):
        for col in range(COLS):
            add('label', f'PlaneCell_{row}_{col}', PLANE['x']+12*col, PLANE['y']+9*row, 12, 9,
                background=hex_color(0, col/(COLS-1), 1-row/(ROWS-1)), paint_cell='plane')
    for col in range(HUE_COLS):
        add('label', f'HueCell_{col}', HUE['x']+6*col, HUE['y'], 6, HUE['h'],
            background=hex_color(1-col/(HUE_COLS-1), 1, 1), paint_cell='hue')
    for name, x in [('Original', 18), ('Current', 96)]:
        add('label', name, x, 278, 56, 26, background='#C00000', color_swatch=True)
        add('label', name+'Label', x, 308, 56, 20, '原色' if name=='Original' else '新色', align='center')
    for index, color in enumerate(SWATCHES):
        add('label', 'Swatch'+str(index), 184+(index%10)*22, 278+(index//10)*24, 20, 20,
            background=color, color_swatch=True, action='swatch', value=index, event='Click')
    for name, caption, x, w in [('Hex', 'HEX', 18, 138), ('R', 'R', 168, 70), ('G', 'G', 250, 70), ('B', 'B', 332, 70)]:
        add('textedit', name, x, 384, w, 28, single_line=True, action='hex' if name=='Hex' else 'rgb', event='Change')
        add('label', name+'Label', x, 416, w, 18, caption, align='center')
    add('label', 'ErrorText', 18, 442, 384, 24, foreground='#B42318')
    add('commandbutton', 'Cancel', 218, 480, 82, 28, '取消', action='cancel', event='Click')
    add('commandbutton', 'Apply', 312, 480, 90, 28, '使用颜色', action='apply', event='Click', primary=True)
    # These must be last in z-order: a single hit surface covers each gradient.
    for name, surface, rect in [('Plane', 'plane', PLANE), ('Hue', 'hue', HUE)]:
        add('label', name, **rect, transparent=True, surface=surface)
    return items


def build_form():
    return build_resource('TableStyleColorForm', '选择颜色', WIDTH, HEIGHT, definitions(), modal=True)


def event_code():
    code = ''
    for item in definitions():
        prefix = 'TableStyleColorForm_'+item['name']
        if 'surface' in item:
            for event, phase in [('MouseDown', 'down'), ('MouseMove', 'move'), ('MouseUp', 'up')]:
                payload = '{surface: '+json.dumps(item['surface'])+', phase: '+json.dumps(phase)+', button: btn, x: x, y: y}'
                code += f'function {prefix}_{event}(btn, shift, x, y) {{ return TableStylesUserFormColorAction("pointer", {payload}); }}\n'
        elif 'action' in item:
            value = str(item['value']) if 'value' in item else 'null'
            code += f'function {prefix}_{item["event"]}() {{ return TableStylesUserFormColorAction({json.dumps(item["action"])}, {value}); }}\n'
            if item['kind'] == 'textedit':
                code += f'function {prefix}_KeyUp(keycode, shift) {{ if (Number(keycode) === 13 || Number(keycode) === 16777220 || Number(keycode) === 16777221) return TableStylesUserFormColorAction("commit"); }}\n'
    return code
