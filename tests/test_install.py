"""Exercise offline install transactions and platform paths in isolated directories."""
from contextlib import redirect_stdout
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from xml.etree import ElementTree as ET

spec = importlib.util.spec_from_file_location('addin_installer', Path(__file__).resolve().parents[1] / 'scripts/install.py')
installer = importlib.util.module_from_spec(spec)
spec.loader.exec_module(installer)


class OfflineInstallTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix='wps-offline-install-')
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.source = self.root / '解压目录 with spaces' / 'dist'
        self.source.mkdir(parents=True)
        for name in installer.REQUIRED:
            (self.source / name).parent.mkdir(parents=True, exist_ok=True)
            (self.source / name).write_text('resource: ' + name, encoding='utf-8')
        (self.source / 'index.html').write_text('old online entry', encoding='utf-8')
        self.config = self.root / 'wps 用户目录' / 'jsaddons' / 'publish.xml'
        self.version = '1.1.0'
        self.output = redirect_stdout(io.StringIO())
        self.output.__enter__()
        self.addCleanup(self.output.__exit__, None, None, None)

    def install(self):
        return installer.install(self.config, self.source, self.version)

    def existing_config(self):
        self.config.parent.mkdir(parents=True, exist_ok=True)
        raw = ('<jsplugins><!--keep comment--><jspluginonline name="Other" type="et" url="https://example.com/"/>'
               f'<jspluginonline name="{installer.NAME}" type="wpp" url="http://127.0.0.1:3989/"/>'
               '</jsplugins>')
        self.config.write_text(raw, encoding='utf-8')
        return raw

    def test_installed_resources_are_local_copies_without_online_entry(self):
        destination = self.install()
        node = ET.parse(self.config).getroot().find('jsplugin')
        self.assertEqual(node.get('type'), 'wpp')
        self.assertEqual(node.get('url'), destination.name)
        self.assertNotIn('://', node.get('url'))
        self.assertFalse((destination / 'index.html').exists())
        self.assertEqual(json.loads((destination / installer.MARKER).read_text())['name'], installer.NAME)
        for name in installer.REQUIRED:
            (self.source / name).unlink()
            self.assertEqual((destination / name).read_text(), 'resource: ' + name)

    def test_online_migration_preserves_other_addins_and_backs_up_original_config(self):
        raw = self.existing_config()
        self.install()
        root = ET.parse(self.config).getroot()
        self.assertEqual(len([n for n in root if n.get('name') == installer.NAME]), 1)
        self.assertIsNotNone(root.find("jspluginonline[@name='Other']"))
        self.assertIn('<!--keep comment-->', self.config.read_text())
        self.assertEqual(next(self.config.parent.glob('publish.xml.bak-*')).read_text(), raw)

    def test_reinstall_backs_up_old_resources_and_registers_once(self):
        destination = self.install()
        (destination / 'js/main.js').write_text('previous installation')
        self.install()
        old = next(self.config.parent.glob(destination.name + '.bak-*'))
        self.assertEqual((old / 'js/main.js').read_text(), 'previous installation')
        self.assertEqual((destination / 'js/main.js').read_text(), 'resource: js/main.js')
        self.assertEqual(len(ET.parse(self.config).getroot()), 1)
        self.assertFalse(list(self.config.parent.glob('.' + installer.NAME + '-install-*')))

    def test_uninstall_only_removes_own_registration_and_managed_resources(self):
        self.existing_config()
        destination = self.install()
        unrelated = self.config.parent / (installer.NAME + '_unmanaged')
        unrelated.mkdir()
        (unrelated / 'keep').write_text('user file')
        installer.uninstall(self.config)
        installer.uninstall(self.config)
        self.assertFalse(destination.exists())
        self.assertEqual((unrelated / 'keep').read_text(), 'user file')
        self.assertIsNotNone(ET.parse(self.config).getroot().find("jspluginonline[@name='Other']"))
        self.assertTrue(list(self.config.parent.glob('publish.xml.bak-*')))

    def test_failed_config_commit_restores_previous_resources_and_config(self):
        destination = self.install()
        (destination / 'js/main.js').write_text('previous installation')
        before = self.config.read_bytes()
        with patch.object(installer, 'write_config', side_effect=OSError('simulated commit failure')):
            with self.assertRaisesRegex(OSError, 'commit failure'):
                self.install()
        self.assertEqual(self.config.read_bytes(), before)
        self.assertEqual((destination / 'js/main.js').read_text(), 'previous installation')
        self.assertFalse(list(self.config.parent.glob('.' + installer.NAME + '-install-*')))

    def test_missing_resources_and_invalid_config_do_not_change_existing_files(self):
        before = self.existing_config()
        (self.source / 'js/main.js').unlink()
        with self.assertRaisesRegex(ValueError, 'js/main.js'):
            self.install()
        self.assertEqual(self.config.read_text(), before)
        self.assertEqual(list(self.config.parent.iterdir()), [self.config])
        (self.source / 'js/main.js').write_text('fixed')
        self.config.write_text('<jsplugins>broken')
        with self.assertRaises(ET.ParseError):
            self.install()
        self.assertEqual(self.config.read_text(), '<jsplugins>broken')
        self.assertEqual(list(self.config.parent.iterdir()), [self.config])

    def test_does_not_replace_an_unmanaged_destination(self):
        destination = self.config.parent / f'{installer.NAME}_{self.version}'
        destination.mkdir(parents=True)
        (destination / 'keep').write_text('user file')
        with self.assertRaisesRegex(ValueError, '未覆盖'):
            self.install()
        self.assertEqual((destination / 'keep').read_text(), 'user file')
        self.assertFalse(self.config.exists())

    def test_default_config_paths_match_each_platform_and_prefer_existing_config(self):
        self.assertEqual(installer.default_config('linux', self.root),
                         self.root / '.local/share/Kingsoft/wps/jsaddons/publish.xml')
        self.assertEqual(installer.default_config('darwin', self.root),
                         self.root / 'Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml')
        mac_old = self.root / 'Library/Application Support/Kingsoft/wps/jsaddons/publish.xml'
        mac_old.parent.mkdir(parents=True); mac_old.write_text('<jsplugins/>')
        self.assertEqual(installer.default_config('darwin', self.root), mac_old)
        roaming = self.root / 'Windows AppData'
        self.assertEqual(installer.default_config('win32', self.root, {'APPDATA': str(roaming)}),
                         roaming / 'Kingsoft/wps/jsaddons/publish.xml')
        with self.assertRaisesRegex(ValueError, '不支持'):
            installer.default_config('unknown', self.root)



if __name__ == '__main__':
    unittest.main()
