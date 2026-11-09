"""Validate the released UserForm-only PPTM, including its OPC relationships."""
import importlib.util
from pathlib import Path
import posixpath
import re
import unittest
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('package_userform_macro', ROOT/'scripts/package_userform_macro.py')
package = importlib.util.module_from_spec(spec)
spec.loader.exec_module(package)


class UserFormMacroTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.entries = package.build_entries('var NativeUserFormTableStyles = {};')
        cls.project = ET.fromstring(cls.entries['ppt/JDEData.bin'])

    def test_only_one_static_guide_slide_without_control_parts_or_macro_actions(self):
        slides = [n for n in self.entries if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)]
        self.assertEqual(slides, ['ppt/slides/slide1.xml'])
        self.assertFalse(any(n.startswith(('ppt/jdecontrols/', 'ppt/notesSlides/')) for n in self.entries))
        slide = ET.fromstring(self.entries[slides[0]])
        self.assertFalse(slide.findall(f'.//{{{package.A}}}hlinkClick'))
        self.assertFalse(slide.findall(f'.//{{{package.P}}}pic'))
        self.assertNotIn('ppaction://macro', self.entries[slides[0]].decode())
        presentation = ET.fromstring(self.entries['ppt/presentation.xml'])
        self.assertEqual(len(presentation.find(f'{{{package.P}}}sldIdLst')), 1)
        size = presentation.find(f'{{{package.P}}}sldSz')
        for transform in slide.findall(f'.//{{{package.A}}}xfrm'):
            offset = transform.find(f'{{{package.A}}}off')
            extent = transform.find(f'{{{package.A}}}ext')
            self.assertLessEqual(int(offset.get('x'))+int(extent.get('cx')), int(size.get('cx')))
            self.assertLessEqual(int(offset.get('y'))+int(extent.get('cy')), int(size.get('cy')))

    def test_all_relationship_targets_and_content_overrides_exist(self):
        for name, data in self.entries.items():
            if name.endswith(('.xml', '.rels')): ET.fromstring(data)
            if not name.endswith('.rels'): continue
            directory = posixpath.dirname(posixpath.dirname(name))
            rels = ET.fromstring(data)
            ids = [n.get('Id') for n in rels]
            self.assertEqual(len(ids), len(set(ids)), name)
            for rel in rels:
                if rel.get('TargetMode') == 'External': continue
                target = posixpath.normpath(posixpath.join(directory, rel.get('Target'))).lstrip('/')
                self.assertIn(target, self.entries, name+' → '+target)
        types = ET.fromstring(self.entries['[Content_Types].xml'])
        names = [n.get('PartName') for n in types if n.get('PartName')]
        self.assertEqual(len(names), len(set(names)))
        for name in names: self.assertIn(name.lstrip('/'), self.entries)
        self.assertIn('macroEnabled', next(n.get('ContentType') for n in types if n.get('PartName') == '/ppt/presentation.xml'))

    def test_two_native_forms_and_all_event_handlers_are_embedded(self):
        package.validate_project(self.project)
        self.assertEqual([n.get('name') for n in self.project.findall('dlgmodule')], ['TableStyleForm', 'TableStyleColorForm'])
        self.assertEqual([n.get('name') for n in self.project.findall('codemodule')], ['TableStyleEditor', 'TableStyleRuntime'])
        source = self.project.findtext('codemodule/codetext')
        self.assertNotIn('__TABLE_STYLES_', source)
        self.assertNotIn('TsOpenUserForm', source)
        self.assertNotIn('TableStylesSlideEditor', source)
        self.assertIn('function OpenTableStyleEditor()', source)
        self.assertIn('new runtime.UserFormColorSurface', source)
        cursor = int(self.project.find('codemodule/window').get('cursorpos'))
        self.assertEqual(source[cursor-1], '{')
        for item in package.definitions():
            if 'action' in item: self.assertIn(f'function TableStyleForm_{item["name"]}_{item["event"]}()', source)
        for name in ['Plane', 'Hue']:
            for event in ['MouseDown', 'MouseMove', 'MouseUp']:
                self.assertIn(f'function TableStyleColorForm_{name}_{event}', source)
        form = self.project.find("dlgmodule[@name='TableStyleForm']/control")
        self.assertEqual(len(form.findall('control')), len(package.definitions()))
        palette = self.project.find("dlgmodule[@name='TableStyleColorForm']/control")
        self.assertEqual(palette.findtext("property[@type='facade']/show_modal"), 'true')


if __name__ == '__main__': unittest.main()
