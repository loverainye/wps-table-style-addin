# JS 宏模板 1.0

当前发行文件为 `表格样式编辑器-幻灯片宏版-1.0.pptm`，同一文件包含幻灯片编辑器和原生 UserForm 编辑器。完整说明见 [SLIDE-README.md](SLIDE-README.md)。

打开模板、启用 JS 宏，第 1 页使用幻灯片编辑器，第 4 页打开 UserForm 窗体。宏编辑器也可运行 `OpenTableStyleEditor` 或 `OpenTableStyleUserForm`。颜色按钮打开各自的色盘；UserForm 色盘同样是原生窗体。

使用「打开 PPTX」选择已关闭的目标文件，编辑后点击保存。模板需保留为 PPTM，目标文件仍为普通 PPTX。目标文件旁自动保留备份，保存不自动打开或重开文稿。若目标文件已在 WPS 打开，手动关闭时不要用内存中的旧内容覆盖磁盘上的新样式，再重新打开查看。

宏版需要 JS 宏、UserForm、FileDialog 和文件读写能力。文件读写优先使用 WPS 的全局 `FileSystem` / `Application.FileSystem`，缺少时才调用对应系统的文件接口。UUID 优先使用 JS Crypto，没有时直接用纯 JS 生成，不依赖任何系统文件或 FFI。模板无需按系统分别下载，实际可用性仍取决于该平台 WPS 版本是否开放相应宏能力。详见项目 [README](../README.md)。

`runtime.js`、`form.py` 等文件还保留早期独立窗体的构建入口，当前发布使用 `slide-runtime.js`、`userform-runtime.js` 和 `scripts/package_slide_macro.py`。
