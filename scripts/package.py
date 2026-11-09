#!/usr/bin/env python3
"""Package prebuilt installers for Linux, macOS, and Windows."""
import json
from pathlib import Path
import tarfile
import zipfile

ROOT = Path(__file__).resolve().parent.parent
version = json.loads((ROOT / 'package.json').read_text(encoding='utf-8'))['version']
release = ROOT / 'release'
release.mkdir(exist_ok=True)
required = ['install.sh', 'uninstall.sh', 'README.md', 'package.json', 'scripts/install.py', 'macro/README.md', 'macro/SLIDE-README.md']
for name in required + ['dist/js/main.js', 'dist/ribbon.xml', 'dist/editor.html', 'dist/editor.js', 'dist/editor.css', 'dist/table-style.svg']:
    if not (ROOT / name).is_file():
        raise SystemExit(f'缺少 {name}，请先构建。')
if (ROOT / 'dist/index.html').exists():
    raise SystemExit('离线构建不应包含 index.html，请先运行 npm run build。')


def metadata(info):
    info.uid = info.gid = 0
    info.uname = info.gname = ''
    if info.name.endswith(('/install.sh', '/uninstall.sh')):
        info.mode = 0o755
    return info


def files_for(kind):
    names = required + (['install.ps1', 'uninstall.ps1'] if kind == 'windows' else [])
    return names + [str(path.relative_to(ROOT)) for path in sorted((ROOT / 'dist').rglob('*'))
                    if path.is_file() and path.suffix != '.map']


for kind, extension in [('linux', 'tar.gz'), ('macos', 'tar.gz')]:
    folder = f'wps-table-style-addin-{kind}-{version}'
    output = release / f'{folder}.{extension}'
    with tarfile.open(output, 'w:gz') as archive:
        for name in files_for(kind):
            archive.add(ROOT / name, arcname=f'{folder}/{name}', filter=metadata)
    print(output)

folder = f'wps-table-style-addin-windows-{version}'
output = release / f'{folder}.zip'
with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
    for name in files_for('windows'):
        archive.write(ROOT / name, f'{folder}/{name}')
print(output)
print('跨平台安装包已生成；运行时只需要对应平台的 WPS。')
