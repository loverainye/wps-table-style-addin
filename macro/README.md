# JS 宏模板 1.0

当前发行版仅使用原生 UserForm：`TableStyleForm` 编辑样式，`TableStyleColorForm` 选择颜色。PPTM 内的一页幻灯片仅为静态说明，没有编辑器、色盘、启动按钮或放映控件。完整说明见 [USERFORM-README.md](USERFORM-README.md)。

打开模板并启用 JS 宏，在宏列表或宏编辑器运行 `OpenTableStyleEditor` 或 `OpenTableStyleUserForm`。使用「打开 PPTX」选择目标文件，编辑后保存。模板需保留为 PPTM，目标文件仍为普通 PPTX。

宏直接读写已保存的目标文件；不会生成临时快照、调用 `SaveCopyAs`、保存正文或自动重新打开目标文稿。当前文稿有未保存修改时需先保存正文。目标文件旁仍保留恢复备份。写入样式后手动重开目标文件，避免使用 WPS 内存中的旧内容覆盖磁盘上的新样式。

文件读写优先使用 WPS 的全局 `FileSystem` / `Application.FileSystem`，缺少时才调用对应系统的文件接口。UUID 优先使用 JS Crypto，没有时直接用纯 JS 生成，不依赖系统文件或 FFI。宏版不使用临时目录或临时图片。模板无需按系统分别下载，仍需该平台 WPS 支持 JS 宏、UserForm 和文件读写。详见项目 [README](../README.md)。

当前发布入口为 `userform-entry-runtime.js` 和 `scripts/package_userform_macro.py`。早期幻灯片版本的源码和回归用例仅保留作兼容记录，不进入当前 PPTM 或安装包。
