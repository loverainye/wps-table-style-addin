"""Check native slide bindings against the independently WPS-saved sample."""
import importlib.util
from pathlib import Path
import posixpath
import re
import unittest
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('package_slide_macro', ROOT/'scripts/package_slide_macro.py')
package = importlib.util.module_from_spec(spec)
spec.loader.exec_module(package)


class SlideMacroTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.items = package.definitions()+package.color_definitions(2)+package.color_definitions(3)+package.launcher_definitions()
        cls.source = 'function OpenTableStyleEditor() {}\n'+package.event_code(package.definitions())+'\n'+package.color_events(package.color_definitions(2)+package.color_definitions(3))
        cls.entries = package.build_entries(cls.source, lambda item: b'PNG placeholder in structural test')

    def test_controls_match_native_property_groups_and_picture_binding(self):
        native = ET.parse(ROOT/'tests/fixtures/wps-native-slide-controls.xml').getroot()
        templates = {c.get('type'): c for c in native.findall(f'{{{package.JDE}}}jdecontrol/control')}
        sample_pic = native.find(f'.//{{{package.P}}}pic')
        sample_ref = sample_pic.find(f'{{{package.P}}}spPr/{{{package.A}}}extLst/{{{package.JDE}}}jdecontrolRef')
        self.assertIsNotNone(sample_ref)
        pics = [pic for page in (1, 2, 3, 4) for pic in ET.fromstring(self.entries[f'ppt/slides/slide{page}.xml']).findall(f'.//{{{package.P}}}pic')]
        self.assertEqual(len(pics), len(self.items))
        for number, (item, pic) in enumerate(zip(self.items, pics), 1):
            node = ET.fromstring(self.entries[f'ppt/jdecontrols/jdecontrol{number}.xml']).find('control')
            sample = templates[node.get('type')]
            self.assertEqual([p.get('type') for p in node], [p.get('type') for p in sample], item['name'])
            for group in sample.findall('property'):
                expected = [x.tag for x in group]
                actual = [x.tag for x in node.find(f"property[@type='{group.get('type')}']")]
                self.assertEqual(actual, expected, item['name'])
            self.assertEqual(node.findtext("property[@type='facade']/name"), item['native'])
            self.assertEqual(pic.find(f'.//{{{package.P}}}cNvPr').get('name'), item['native'])
            ref = pic.find(f'{{{package.P}}}spPr/{{{package.A}}}extLst/{{{package.JDE}}}jdecontrolRef')
            self.assertEqual(ref.get(f'{{{package.R}}}id'), f'rIdControl{number}')
            for key, source in [('left', 'x'), ('top', 'y'), ('width', 'w'), ('height', 'h')]:
                value = int(node.findtext(f"property[@type='geometry']/{key}"))
                self.assertEqual(value, item[source]*12700)
            xfrm = pic.find(f'{{{package.P}}}spPr/{{{package.A}}}xfrm')
            for tag, attr, source in [('off', 'x', 'x'), ('off', 'y', 'y'), ('ext', 'cx', 'w'), ('ext', 'cy', 'h')]:
                self.assertEqual(int(xfrm.find(f'{{{package.A}}}{tag}').get(attr)), item[source]*12700)

    def test_events_have_unique_document_names_and_no_userform_prefix(self):
        names = [x['native'] for x in self.items]
        self.assertEqual(len(names), len(set(names)))
        self.assertNotIn('TableStyleForm_', self.source)
        self.assertIn('function TsOpenFile_Click()', self.source)
        self.assertIn('function TsFill_Change()', self.source)
        for item in self.items:
            if item.get('page') == 4: continue
            self.assertIn(f"{item['name']}: {item['native']}", self.source)
            if 'action' in item:
                self.assertIn(f"function {item['native']}_{item['event']}()", self.source)
        project = ET.fromstring(self.entries['ppt/JDEData.bin'])
        self.assertEqual(project.find('dlgmodule').get('name'), 'TableStyleForm')
        package.validate_project(project)
        self.assertEqual(len(project.findall('codemodule')), 4)
        self.assertEqual([x.tag for x in project.find('codemodule')], ['window', 'codetext'])

    def test_default_editor_module_stays_small_and_runtime_has_no_module_level_initializers(self):
        source = (ROOT/'macro/slide-code.js').read_text()
        project = package.build_project(source, 'function TableStylesNativeRuntime() {}', self.source)
        modules = project.findall('codemodule')
        self.assertLess(len(modules[0].findtext('codetext')), 10000)
        self.assertEqual(modules[0].get('name'), 'TableStyleEditor')
        self.assertEqual(modules[0].find('window').get('actived'), 'true')
        self.assertEqual(modules[2].get('name'), 'TableStyleRuntime')
        self.assertEqual(modules[2].find('window').get('visible'), 'false')
        self.assertTrue(modules[2].findtext('codetext').startswith('function TableStylesNativeRuntime()'))

    def test_colors_use_one_dropdown_and_palette_button_including_every_border(self):
        items = {item['name']: item for item in self.items if not item.get('picker')}
        self.assertNotIn('FillChoice', items)
        self.assertNotIn('TextColorChoice', items)
        for name in ['Fill', 'TextColor', 'Color_left', 'Color_right', 'Color_top', 'Color_bottom', 'Color_insideH', 'Color_insideV']:
            self.assertEqual(items[name]['kind'], 'combobox')
            self.assertEqual(items['Pick'+name]['kind'], 'label')
            self.assertEqual(items['Pick'+name]['caption'], '\u00a0')
            node = ET.fromstring(package.control_part(items['Pick'+name])).find('control')
            self.assertEqual(node.findtext("property[@type='facade']/back_style"), 'true')
            self.assertEqual(node.findtext("property[@type='facade']/border_style"), 'true')
            self.assertEqual(items['Pick'+name]['action'], 'pick:'+name)

    def test_palette_has_transparent_pointer_surfaces_and_stays_visible_on_its_own_page(self):
        items = package.color_definitions(2)+package.color_definitions(3)
        events = package.color_events(items)
        self.assertEqual(len(items), 38)
        for item in items:
            node = ET.fromstring(package.control_part(item)).find('control')
            self.assertEqual(node.findtext("property[@type='facade']/visible"), 'true')
            if item['kind'] == 'label':
                self.assertTrue(node.findtext("property[@type='facade']/caption"), 'empty caption can retain WPS Label1 default')
            if item.get('surface'):
                self.assertEqual(node.findtext("property[@type='facade']/back_style"), 'false')
                for event in ['MouseDown', 'MouseMove', 'MouseUp']:
                    self.assertIn(f'function {item["native"]}_{event}(btn, shift, x, y)', events)
            elif item.get('action'):
                self.assertIn(f'function {item["native"]}_{item["event"]}()', events)
            if item['kind'] == 'textedit':
                self.assertIn(f'function {item["native"]}_KeyUp(keycode, shift)', events)

    def test_picker_shapes_are_prebuilt_on_dedicated_pages_and_below_pointer_controls(self):
        slide = ET.fromstring(self.entries['ppt/slides/slide2.xml'])
        tree = slide.find(f'{{{package.P}}}cSld/{{{package.P}}}spTree')
        nodes = list(tree)
        names = {node.find(f'.//{{{package.P}}}cNvPr').get('name'): node for node in nodes
                 if node.find(f'.//{{{package.P}}}cNvPr') is not None}
        expected = ['TsColorPlaneImage', 'TsColorHueImage']
        first_control = next(i for i, node in enumerate(nodes) if node.tag == f'{{{package.P}}}pic')
        for name in expected:
            node = names[name]
            self.assertEqual(node.tag, f'{{{package.P}}}sp')
            self.assertIsNone(node.find(f'.//{{{package.P}}}cNvPr').get('hidden'))
            self.assertLess(nodes.index(node), first_control)
            geometry = node.find(f'{{{package.P}}}spPr/{{{package.A}}}prstGeom').get('prst')
            self.assertEqual(geometry, 'rect' if name.endswith('Image') else 'ellipse')
        for name, geometry in [('TsColorPlaneImage', package.PLANE), ('TsColorHueImage', package.HUE)]:
            transform = names[name].find(f'{{{package.P}}}spPr/{{{package.A}}}xfrm')
            for tag, attr, key in [('off', 'x', 'x'), ('off', 'y', 'y'), ('ext', 'cx', 'w'), ('ext', 'cy', 'h')]:
                self.assertEqual(int(transform.find(f'{{{package.A}}}{tag}').get(attr)), geometry[key]*package.EMU)
        ids = [node.get('id') for node in tree.iter(f'{{{package.P}}}cNvPr')]
        self.assertEqual(len(ids), len(set(ids)))
        self.assertEqual(len(self.items), 153)
        for page, count in [(1, 110), (2, 19), (3, 19), (4, 5)]:
            root = ET.fromstring(self.entries[f'ppt/slides/slide{page}.xml'])
            self.assertEqual(len(root.findall(f'.//{{{package.P}}}pic')), count)
            page_names = [n.get('name') for n in root.iter(f'{{{package.P}}}cNvPr')]
            self.assertEqual(len(set(expected).intersection(page_names)), 2 if page in (2, 3) else 0)
            page_ids = [n.get('id') for n in root.iter(f'{{{package.P}}}cNvPr')]
            self.assertEqual(len(page_ids), len(set(page_ids)))
            self.assertEqual(root.find(f'{{{package.P}}}transition').get('advClick'), '0')

    def test_relationships_types_and_all_four_slides_are_complete(self):
        for path, data in self.entries.items():
            if not path.endswith('.rels'): continue
            directory = posixpath.dirname(posixpath.dirname(path))
            rels = ET.fromstring(data)
            ids = [x.get('Id') for x in rels]
            self.assertEqual(len(ids), len(set(ids)), path)
            for rel in rels:
                if rel.get('TargetMode') == 'External': continue
                target = posixpath.normpath(posixpath.join(directory, rel.get('Target')))
                if target.startswith('/'): target = target[1:]
                self.assertIn(target, self.entries, path+' → '+target)
        types = ET.fromstring(self.entries['[Content_Types].xml'])
        overrides = {x.get('PartName'): x.get('ContentType') for x in types if x.get('PartName')}
        self.assertIn('macroEnabled', overrides['/ppt/presentation.xml'])
        for number in range(1, len(self.items)+1):
            self.assertEqual(overrides[f'/ppt/jdecontrols/jdecontrol{number}.xml'],
                             'application/vnd.wps-officedocument.jdeControl+xml')
        slides = [name for name in self.entries if re.fullmatch('ppt/slides/slide[0-9]+.xml', name)]
        self.assertEqual(slides, [f'ppt/slides/slide{page}.xml' for page in (1, 2, 3, 4)])
        presentation = ET.fromstring(self.entries['ppt/presentation.xml'])
        self.assertEqual(len(presentation.find(f'{{{package.P}}}sldIdLst')), 4)

    def test_userform_resource_events_and_launcher_are_in_the_same_project(self):
        project = ET.fromstring(self.entries['ppt/JDEData.bin'])
        ids = [node.get('id') for node in project if node.tag in ('codemodule', 'dlgmodule')]
        self.assertEqual(len(ids), len(set(ids)))
        dialog = project.find('dlgmodule')
        self.assertEqual([node.tag for node in dialog], ['window', 'control'])
        self.assertEqual(dialog.find('window').get('visible'), 'false')
        form = dialog.find('control')
        sample = ET.parse(ROOT/'tests/fixtures/wps-native-userform.xml').find('dlgmodule/control')
        self.assertEqual([p.get('type') for p in form.findall('property')],
                         [p.get('type') for p in sample.findall('property')])
        self.assertEqual(form.findtext("property[@type='geometry']/width"), str(700*12700))
        source = project.find("codemodule[@name='TableStyleUserForm']/codetext").text
        self.assertNotIn('__TABLE_STYLES_USERFORM_GEOMETRY__', source)
        self.assertIn('function TsOpenUserForm_Click() { return OpenTableStyleUserForm(); }', source)
        self.assertIn('TableStyleForm.Show()', source)
        self.assertNotIn('TsOpenUserForm', self.source)
        items = package.userform_definitions()
        for item, node in zip(items, form.findall('control')):
            self.assertEqual(node.findtext("property[@type='facade']/name"), item['name'])
            if 'action' in item:
                self.assertIn(f'function TableStyleForm_{item["name"]}_{item["event"]}()', source)
            for key, source_key in [('left', 'x'), ('top', 'y'), ('width', 'w'), ('height', 'h')]:
                self.assertEqual(int(node.findtext(f"property[@type='geometry']/{key}")), item[source_key]*12700)

    def test_userform_fixed_layout_and_step_buttons_fit_without_overlaps(self):
        items = package.userform_definitions()
        names = [i['name'] for i in items]
        self.assertEqual(len(names), len(set(names)))
        self.assertFalse(set(['ZoomIn', 'ZoomOut', 'ZoomLevel']).intersection(names))
        self.assertEqual(len([n for n in names if n.startswith('Step')]), 14)
        for item in items:
            self.assertGreaterEqual(item['x'], 0)
            self.assertGreaterEqual(item['y'], 0)
            self.assertLessEqual(item['x']+item['w'], 700)
            self.assertLessEqual(item['y']+item['h'], 762)
        # Labels have a deliberate baseline offset; their bounds must also fit
        # between rows without covering the actual interactive controls.
        for i, first in enumerate(items):
            for second in items[i+1:]:
                overlap = (min(first['x']+first['w'], second['x']+second['w']) > max(first['x'], second['x'])
                           and min(first['y']+first['h'], second['y']+second['h']) > max(first['y'], second['y']))
                self.assertFalse(overlap, first['name']+' overlaps '+second['name'])

    def test_userform_colors_are_single_native_dropdowns_with_palette_swatches(self):
        root, items = package.build_userform()
        controls = {i['name']: (i, node) for i, node in zip(items, root.findall('control'))}
        self.assertNotIn('FillChoice', controls)
        self.assertNotIn('TextColorChoice', controls)
        for name in ['Fill', 'TextColor', 'Color_left', 'Color_right', 'Color_top', 'Color_bottom', 'Color_insideH', 'Color_insideV']:
            item, node = controls[name]
            self.assertEqual(item['kind'], 'combobox')
            self.assertEqual(node.get('type'), 'combobox')
            self.assertEqual(node.findtext("property[@type='data']/list_rows"), '8')
            self.assertIsNotNone(node.find("property[@type='facade']/value"))
            swatch, resource = controls['Pick'+name]
            self.assertEqual(swatch['action'], 'pick:'+name)
            self.assertEqual(swatch['caption'], '\u00a0')
            self.assertEqual(resource.findtext("property[@type='facade']/border_style"), 'true')
            self.assertEqual(resource.findtext("property[@type='facade']/back_style"), 'true')
            self.assertGreaterEqual(swatch['x'], item['x']+item['w'])
        # Text fields use the full native text/value schema rather than defaults.
        width = controls['Width_left'][1]
        self.assertEqual(width.get('type'), 'textedit')
        self.assertIsNotNone(width.find("property[@type='facade']/value"))
        self.assertEqual(width.findtext("property[@type='actions']/word_wrap"), 'false')

    def test_native_color_form_contains_its_own_gradients_inputs_and_mouse_surfaces(self):
        from userform_color_picker import definitions, PLANE, HUE, COLS, ROWS, HUE_COLS, WIDTH, HEIGHT
        project = ET.fromstring(self.entries['ppt/JDEData.bin'])
        self.assertEqual(len(project.findall('dlgmodule')), 2)
        dialog = project.find("dlgmodule[@name='TableStyleColorForm']")
        self.assertEqual([n.tag for n in dialog], ['window', 'control'])
        form = dialog.find('control')
        self.assertEqual(form.findtext("property[@type='facade']/show_modal"), 'true')
        self.assertEqual(form.findtext("property[@type='facade']/caption"), '选择颜色')
        items = definitions()
        self.assertEqual(len(form.findall('control')), len(items))
        names = [i['name'] for i in items]
        self.assertEqual(len(names), len(set(names)))
        self.assertEqual(len([i for i in items if i.get('paint_cell') == 'plane']), COLS*ROWS)
        self.assertEqual(len([i for i in items if i.get('paint_cell') == 'hue']), HUE_COLS)
        self.assertEqual(names[-2:], ['Plane', 'Hue'], 'mouse surfaces must be above the color cells')
        source = project.find("codemodule[@name='TableStyleUserForm']/codetext").text
        self.assertNotIn('SlideColorSurface', source)
        self.assertNotIn('TableStylesPickerControls', source)
        self.assertIn('new runtime.UserFormColorSurface(TableStyleColorForm', source)
        self.assertLess(len(source), 50000, 'color cells must not generate a function each')
        for name in ['Plane', 'Hue']:
            for event in ['MouseDown', 'MouseMove', 'MouseUp']:
                self.assertIn(f'function TableStyleColorForm_{name}_{event}', source)
        for item, node in zip(items, form.findall('control')):
            self.assertGreaterEqual(item['x'], 0)
            self.assertGreaterEqual(item['y'], 0)
            self.assertLessEqual(item['x']+item['w'], WIDTH)
            self.assertLessEqual(item['y']+item['h'], HEIGHT)
            if item.get('surface'):
                self.assertEqual(node.findtext("property[@type='facade']/back_style"), 'false')
            elif item.get('paint_cell'):
                self.assertEqual(node.findtext("property[@type='facade']/back_style"), 'true')
                self.assertEqual(node.findtext("property[@type='facade']/border_style"), 'false')
        # Non-layered UI controls must not overlap; painted cells deliberately
        # sit under their pointer surface and exactly tile their plane/strip.
        ui = [i for i in items if not i.get('paint_cell') and not i.get('surface')]
        for i, first in enumerate(ui):
            for second in ui[i+1:]:
                overlap = (min(first['x']+first['w'], second['x']+second['w']) > max(first['x'], second['x'])
                           and min(first['y']+first['h'], second['y']+second['h']) > max(first['y'], second['y']))
                self.assertFalse(overlap, first['name']+' overlaps '+second['name'])
        for key, rect in [('plane', PLANE), ('hue', HUE)]:
            cells = [i for i in items if i.get('paint_cell') == key]
            self.assertEqual(sum(i['w']*i['h'] for i in cells), rect['w']*rect['h'])

    def test_layout_fits_page_and_controls_do_not_overlap(self):
        for item in self.items:
            self.assertGreater(item['w'], 0)
            self.assertGreater(item['h'], 0)
            self.assertGreaterEqual(item['x'], 0)
            self.assertGreaterEqual(item['y'], 0)
            self.assertLessEqual(item['x']+item['w'], package.WIDTH)
            self.assertLessEqual(item['y']+item['h'], package.HEIGHT)
        for i, first in enumerate(self.items):
            for second in self.items[i+1:]:
                if first.get('page', 1) != second.get('page', 1): continue
                overlap = (min(first['x']+first['w'], second['x']+second['w']) > max(first['x'], second['x'])
                           and min(first['y']+first['h'], second['y']+second['h']) > max(first['y'], second['y']))
                self.assertFalse(overlap, first['name']+' overlaps '+second['name'])


if __name__ == '__main__': unittest.main()
