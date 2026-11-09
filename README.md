# WPS 演示表格样式编辑器

当前版本：**1.0**。项目提供 WPS JS 加载项（直接编辑当前 PPTX）和内嵌 JS 宏模板（无需安装加载项，可打开其它 PPTX）。JS 加载项提供 Linux、macOS、Windows 安装脚本和离线包；JS 宏模板目前面向 Linux 桌面 WPS。

## 安装

### 在线安装（GitHub Raw 脚本）

在线安装脚本从 GitHub Releases 下载固定版本资产，不需要本地 Node.js。Linux/macOS 安装需要 Python 3.9+（仅标准库）；Windows 使用系统自带 PowerShell 5.1+，无需 Python。运行前请先关闭 WPS，安装完成后重新打开 WPS 演示。

Linux：

```bash
curl -fsSL https://raw.githubusercontent.com/loverainye/wps-table-style-addin/main/install-online.sh | bash
```

macOS：

```bash
curl -fsSL https://raw.githubusercontent.com/loverainye/wps-table-style-addin/main/install-online.sh | bash
```

Windows PowerShell：

```powershell
irm https://raw.githubusercontent.com/loverainye/wps-table-style-addin/main/install-online.ps1 | iex
```

在线脚本下载的 JS 加载项资产为 `wps-table-style-addin-<platform>-1.0`，下载地址来自：
`https://github.com/loverainye/wps-table-style-addin/releases/download/v1.0/`。

### 离线安装

从 [GitHub Releases](https://github.com/loverainye/wps-table-style-addin/releases/tag/v1.0) 下载对应平台的 JS 加载项压缩包：

- Linux：`wps-table-style-addin-linux-1.0.tar.gz`
- macOS：`wps-table-style-addin-macos-1.0.tar.gz`
- Windows：`wps-table-style-addin-windows-1.0.zip`

Linux/macOS：

```bash
tar -xzf wps-table-style-addin-<platform>-1.0.tar.gz
cd wps-table-style-addin-<platform>-1.0
bash install.sh
```

将 `<platform>` 替换为 `linux` 或 `macos`。离线安装过程不访问网络；请事先准备好 Python。

Windows PowerShell：

```powershell
Expand-Archive .\wps-table-style-addin-windows-1.0.zip -DestinationPath .
cd .\wps-table-style-addin-windows-1.0
Set-ExecutionPolicy -Scope Process Bypass
.\install.ps1
```

卸载：Linux/macOS 运行 `bash uninstall.sh`；Windows 运行 `.\uninstall.ps1`。安装只复制本地资源和更新 WPS 的 `publish.xml`，不启动 Web 服务。WPS 读取到新配置后需要重新启动 WPS 演示。

### JS 宏模板

Release 同时提供 `表格样式编辑器-幻灯片宏版-1.0.pptm` 和 `wps-table-style-slide-macro-1.0.zip`。打开 PPTM 并启用 JS 宏：

1. 第 1 页是幻灯片版编辑器。
2. 第 4 页点击「打开 UserForm 窗体」，使用原生 `TableStyleForm`。
3. UserForm 的颜色按钮打开原生 `TableStyleColorForm`，取色过程不切换 PPT 页面。
4. 也可以在宏编辑器运行 `OpenTableStyleEditor` 或 `OpenTableStyleUserForm`。

宏模板本身必须保存为 PPTM；它编辑的目标文稿仍然是普通 PPTX。宏模板不需要安装 JS 加载项。

## 功能

- 13 个表格样式区域、填充、文字色、加粗/斜体和六方向边框。
- 边框线宽默认 1 磅，范围 0.25–20 磅，箭头每次 0.25 磅。
- 新建、复制、默认样式、应用范围、默认字号和备份恢复。
- UserForm 颜色字段使用单一颜色下拉框和色块按钮；色盘支持渐变区、色相条、常用色和 HEX/RGB。
- 保存修改 PPTX 内的 tableStyles.xml 及相关引用，不自动关闭或重开目标文件。

保存后需手动重新打开目标文稿才能看到新样式。如果目标文件仍在 WPS 中打开，请在重新打开前避免使用 WPS 原生保存覆盖磁盘上的新样式；关闭时若询问保存，选择不保存。原文件旁会自动生成表格样式备份。宏版推荐打开已关闭的 PPTX。

## 从源码构建

构建需要 Node.js 22、Python 3.9+、Pillow、fontconfig 和中文字体（例如 Noto Sans CJK）。

```bash
npm ci
npm test
npm run build
python3 scripts/package.py
npm run build:slide-macro
python3 scripts/package_slide_macro.py
```

发布工作流位于 `.github/workflows/release.yml`，推送 `v1.0` 标签会运行测试、构建三平台 JS 加载项安装包并发布 JS 宏模板。运行时不需要 Node.js、Python 或网络；在线安装脚本只用于下载 Release 资产。

## 平台说明

Linux 使用 `~/.local/share/Kingsoft/wps/jsaddons`；macOS 默认使用 `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons`，兼容已有 `~/Library/Application Support/Kingsoft/WPS Office/jsaddons` 或 `Kingsoft/wps/jsaddons` 配置；Windows 使用 `%APPDATA%\Kingsoft\wps\jsaddons`，兼容已有 `Kingsoft\WPS Office\jsaddons` 配置。脚本优先使用已经存在的 `publish.xml`，也支持通过 `--config`（PowerShell 为 `-Config`）指定配置文件。

加载项面向具备 JS 加载项、文件读写和任务窗格接口的 WPS 演示客户端，不适用于 WPS 在线文档；macOS 安装版还需允许 WPS 访问上述沙盒目录。安装脚本测试覆盖三平台，但尚未完成三平台 WPS 界面实机验证。

宏模板依赖 JS 宏、原生 UserForm 和 Linux 文件/系统随机 UUID 接口，目前不支持直接在 Windows/macOS 运行。三平台安装支持指 JS 加载项；这些平台的 WPS 版本和授权也需支持 JS 加载项。更多使用说明见 [宏版说明](macro/SLIDE-README.md)。
