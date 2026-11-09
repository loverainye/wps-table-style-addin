# 文件接口说明

当前纯 UserForm 宏只依赖 WPS 原生文件接口，不再包含 Linux、macOS 或 Windows 专用文件实现。

`macro/platform-services.js` 只做两件事：

- 从 `Application.FileSystem` 或宏全局 `FileSystem` 获取 WPS 文件对象；
- 提供 JS UUID 生成器：`crypto.randomUUID()` → `crypto.getRandomValues()` → `Math.random()`。

宏必须具备：

```js
FileSystem.readAsBinaryString(path)
FileSystem.writeAsBinaryString(path, data)
```

如果两个接口不存在，宏会停止并提示使用支持 WPS 原生二进制文件接口的版本。宏不会检测 `Platform.OS()`，不会读取 `/proc`，不会加载 libc、libSystem 或 kernel32，也不会访问临时目录。

`macro/file-host.js` 直接读取和写入已保存的 PPTX。保存恢复备份仍通过同一个 WPS `FileSystem` 写入目标文件旁；没有 `SaveCopyAs`、临时快照或临时图片。

测试使用内存文件对象和本机磁盘文件模拟 WPS `FileSystem` 契约，验证二进制内容、中文路径、备份及错误处理。测试环境不代表 WPS 客户端本身，实际使用仍需 WPS 开放上述接口。
