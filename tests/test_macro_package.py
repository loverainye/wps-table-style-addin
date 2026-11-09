"""Regression checks for the native dialog load crash in macro 1.8.0."""
import importlib.util
from pathlib import Path
import unittest
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('package_macro', ROOT / 'scripts' / 'package_macro.py')
package = importlib.util.module_from_spec(spec)
spec.loader.exec_module(package)


class MacroPackageTests(unittest.TestCase):
    def test_geometry_uses_native_wps_emu_units_instead_of_tiny_point_values(self):
        sample = ET.parse(ROOT / 'tests' / 'fixtures' / 'wps-native-userform.xml')
        native = sample.find('dlgmodule/control')
        # WPS's default form is 375 pt; the button is 90 x 30 pt.
        self.assertEqual(native.findtext("property[@type='geometry']/width"), '4762500')
        button = native.find('control')
        self.assertEqual(button.findtext("property[@type='geometry']/width"), '1143000')
        self.assertEqual(button.findtext("property[@type='geometry']/height"), '381000')
        form, items = package.build_form()
        self.assertEqual(form.findtext("property[@type='geometry']/width"), str(700 * 12700))
        self.assertEqual(form.findtext("property[@type='geometry']/height"), str(762 * 12700))
        for child, item in zip(form.findall('control'), items):
            for key, source in [('left', 'x'), ('top', 'y'), ('width', 'w'), ('height', 'h')]:
                self.assertEqual(int(child.findtext(f"property[@type='geometry']/{key}")),
                                 item[source] * 12700, item['name'])

    def test_native_form_property_groups_and_control_visibility_match_wps_sample(self):
        native = ET.parse(ROOT / 'tests' / 'fixtures' / 'wps-native-userform.xml').find('dlgmodule/control')
        form, _ = package.build_form()
        self.assertEqual([p.get('type') for p in form.findall('property')],
                         [p.get('type') for p in native.findall('property')])
        self.assertEqual(form.findtext("property[@type='layout']/layout_type"), '1')
        self.assertEqual(form.findtext("property[@type='facade']/startup_position"), '1')
        self.assertEqual(form.findtext("property[@type='facade']/show_modal"), 'false')
        self.assertIsNone(form.find("property[@type='others']"))
        for child in form.findall('control'):
            self.assertEqual(child.findtext("property[@type='facade']/visible"), 'true')
            self.assertEqual(child.findtext("property[@type='facade']/back_style"), 'true')

    def test_launcher_runtime_and_events_stay_in_code_module(self):
        form, controls = package.build_form()
        events = package.event_code(controls)
        project = package.build_project('function OpenTableStyleEditor() {}', events, form)
        project = ET.fromstring(ET.tostring(project))
        package.validate_project(project)
        source = project.findtext('codemodule/codetext')
        self.assertIn('function OpenTableStyleEditor()', source)
        self.assertIn('function TableStyleForm_OpenFile_Click()', source)
        self.assertEqual([c.tag for c in project.find('dlgmodule')], ['window', 'control'])
        cursor = int(project.find('codemodule/window').get('cursorpos'))
        self.assertEqual(source[cursor - 1], '{', 'F5 的初始光标应在打开编辑器的函数内')

    def test_rejects_180_code_node_that_crashes_native_control_importer(self):
        form, _ = package.build_form()
        project = package.build_project('', '', form)
        dialog = project.find('dlgmodule')
        ET.SubElement(dialog, 'codetext').text = 'function OpenFile_Click() {}'
        with self.assertRaisesRegex(ValueError, 'codemodule'):
            package.validate_project(project)

    def test_rejects_code_hidden_inside_native_control_resources(self):
        form, _ = package.build_form()
        project = package.build_project('', '', form)
        ET.SubElement(form.find('control'), 'codetext').text = 'function Unexpected() {}'
        with self.assertRaisesRegex(ValueError, '宏代码'):
            package.validate_project(project)


if __name__ == '__main__':
    unittest.main()
