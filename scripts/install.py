#!/usr/bin/env python3
"""Install the WPS add-in on Linux, macOS, or Windows without a web server."""
import argparse
from datetime import datetime
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent.parent
NAME = 'WpsPresentationTableStyles'
MARKER = '.wps-table-style-install.json'
REQUIRED = ('js/main.js', 'ribbon.xml', 'editor.html', 'editor.js', 'editor.css', 'table-style.svg')


def stamp():
    return datetime.now().strftime('%Y%m%d-%H%M%S-%f')


def read_config(config):
    if not config.exists():
        return ET.ElementTree(ET.Element('jsplugins'))
    raw = config.read_bytes()
    if b'<!DOCTYPE' in raw.upper() or b'<!ENTITY' in raw.upper():
        raise ValueError('publish.xml 包含不支持的声明，未修改。')
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))
    tree = ET.ElementTree(ET.fromstring(raw, parser=parser))
    if tree.getroot().tag != 'jsplugins':
        raise ValueError('已有 publish.xml 根节点未知，未修改。')
    return tree


def backup_config(config):
    if config.exists():
        backup = config.with_name(config.name + '.bak-' + stamp())
        shutil.copy2(config, backup)
        print('配置备份：', backup)


def write_config(tree, config):
    ET.indent(tree, space='  ')
    write_bytes(ET.tostring(tree.getroot(), encoding='utf-8', xml_declaration=True), config)


def write_bytes(data, config):
    fd, tmp = tempfile.mkstemp(prefix='.publish-', suffix='.xml.tmp', dir=config.parent)
    try:
        with os.fdopen(fd, 'wb') as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(tmp, config)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


def owned(directory):
    try:
        if directory.is_symlink() or not directory.is_dir():
            return False
        return json.loads((directory / MARKER).read_text(encoding='utf-8')).get('name') == NAME
    except (OSError, ValueError, TypeError, AttributeError):
        return False


def install(config, source=ROOT / 'dist', version=None):
    if version is None:
        version = json.loads((ROOT / 'package.json').read_text(encoding='utf-8'))['version']
    # The version is used as a single local directory component.
    if not version or any(c not in '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.-' for c in version):
        raise ValueError('加载项版本号格式错误。')
    source = Path(source).resolve()
    for filename in REQUIRED:
        if not (source / filename).is_file():
            raise ValueError(f'缺少构建资源 {filename}。请使用完整离线安装包，或先运行 npm run build。')
    tree = read_config(config)
    root = tree.getroot()
    for child in list(root):
        if child.get('name') == NAME:
            root.remove(child)
    folder = f'{NAME}_{version}'
    destination = config.parent / folder
    if destination.is_symlink() or (destination.exists() and not owned(destination)):
        raise ValueError(f'目标目录不是本安装脚本管理的目录，未覆盖：{destination}')
    ET.SubElement(root, 'jsplugin', {
        'name': NAME, 'type': 'wpp', 'url': folder, 'version': version,
        'debug': '', 'enable': 'enable_dev', 'install': 'null', 'customDomain': '',
    })
    config.parent.mkdir(parents=True, exist_ok=True)
    staged = Path(tempfile.mkdtemp(prefix=f'.{NAME}-install-', dir=config.parent))
    previous = None
    moved = False
    try:
        # Linux WPS creates index.html with a script reference to js/main.js.
        # Do not install an online-mode index.html or a cached generated entry.
        shutil.copytree(source, staged, dirs_exist_ok=True,
                        ignore=shutil.ignore_patterns('index.html', '*.map', '__pycache__'))
        (staged / MARKER).write_text(json.dumps({'name': NAME, 'version': version}, ensure_ascii=False) + '\n', encoding='utf-8')
        backup_config(config)
        if destination.exists():
            previous = destination.with_name(destination.name + '.bak-' + stamp())
            destination.rename(previous)
        staged.rename(destination)
        moved = True
        write_config(tree, config)
    except BaseException:
        if moved:
            shutil.rmtree(destination)
        if previous is not None:
            previous.rename(destination)
        raise
    finally:
        if staged.exists():
            shutil.rmtree(staged)
    if previous is not None:
        print('旧资源备份：', previous)
    print(f'已安装 {sys.platform} 离线加载项：', destination)
    print('已注册：', config)
    print('无需 Web 服务。请保存工作并重新打开 WPS 演示，点击“表格样式编辑器 → 编辑样式”。')
    return destination


def uninstall(config):
    tree = read_config(config)
    root = tree.getroot()
    entries = [child for child in root if child.get('name') == NAME]
    for child in entries:
        root.remove(child)
    if entries:
        backup_config(config)
        write_config(tree, config)
    if config.parent.is_dir():
        for directory in config.parent.glob(NAME + '_*'):
            if '.bak-' not in directory.name and owned(directory):
                shutil.rmtree(directory)
                print('已删除加载项资源：', directory)
    print('已移除本加载项；保留其他加载项和备份。请重新打开 WPS 演示。')


def default_config(platform=None, home=None, environ=None):
    platform = platform or sys.platform
    home = Path(home) if home is not None else Path.home()
    environ = os.environ if environ is None else environ
    if platform.startswith('linux'):
        return home / '.local/share/Kingsoft/wps/jsaddons/publish.xml'
    if platform == 'darwin':
        candidates = [home / 'Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml',
                      home / 'Library/Application Support/Kingsoft/WPS Office/jsaddons/publish.xml',
                      home / 'Library/Application Support/Kingsoft/wps/jsaddons/publish.xml']
    elif platform.startswith('win'):
        appdata = Path(environ.get('APPDATA', home / 'AppData/Roaming'))
        candidates = [appdata / 'Kingsoft/wps/jsaddons/publish.xml',
                      appdata / 'Kingsoft/WPS Office/jsaddons/publish.xml']
    else:
        raise ValueError(f'不支持的操作系统：{platform}')
    return next((p for p in candidates if p.exists()), candidates[0])


def main(argv=None):
    parser = argparse.ArgumentParser(description='WPS 演示表格样式编辑器：跨平台离线安装/卸载，无需 sudo。')
    parser.add_argument('--remove', action='store_true', help='卸载本加载项，保留备份及其他加载项')
    parser.add_argument('--config', type=Path, help='指定 publish.xml 路径（测试或自定义 WPS 用户目录）')
    args = parser.parse_args(argv)
    try:
        config = (args.config or default_config()).expanduser().absolute()
        if args.remove:
            uninstall(config)
        else:
            install(config)
    except (OSError, ValueError, ET.ParseError, KeyError) as error:
        print('安装/卸载未完成：' + str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
