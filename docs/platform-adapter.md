# 文件与 UUID 适配层

`macro/platform-services.js` 向编辑器提供 `files`、`tempPath()`、`uuid()`，并通过 `macroApplication()` 包装 WPS 宿主。不会向 WPS 的 Application 代理写入属性。

WPS 文件能力优先，UUID 完全由 JS 生成，各项能力独立选择：

| 能力 | 优先接口 | 缺少时的实现 |
| --- | --- | --- |
| 二进制文件读写 | `Application.FileSystem`、宏全局 `FileSystem` | 当前系统的 FFI 文件接口 |
| 临时目录 | `Application.Env.GetTempPath()`、全局 `Env`、`FileSystem.tmpdir()` | Linux `/tmp`；macOS `TMPDIR`；Windows `GetTempPathW` |
| UUID v4 | JS `crypto.randomUUID()`，其次 `crypto.getRandomValues()` | 纯 JS `Math.random()` 生成 UUID v4 格式的 ID |

`Platform.OS()` 用于判断文件接口所用系统；缺少该接口且缺少原生文件能力时，依次探测可加载的系统库。不会在已识别的 Windows/macOS 上尝试 Linux 文件接口。新系统实现只需遵守适配层契约，编辑器和 ZIP/XML 逻辑无需改动。

| 系统 | 文件实现 |
| --- | --- |
| Linux | `macro/linux-files.js` 加载 libc，复用 POSIX 分块读写 |
| macOS | `macro/platform/macos.js` 加载 `/usr/lib/libSystem.B.dylib`；创建使用固定参数 `creat`，避开 Apple Silicon 可变参数 ABI |
| Windows | `macro/platform/windows.js` 调用 kernel32 的 Unicode 文件接口，支持中文、盘符和 UNC 路径 |

`src/uuid.js` 统一输出带花括号的大写 UUID v4 格式 ID，顺序为 `crypto.randomUUID()` → `crypto.getRandomValues()` → `Math.random()`。宏和加载项共用实现，无需读取系统 UUID 文件、加载系统随机数库或启动外部程序。

FFI 为 WPS 宏自带接口，未随模板附带 DLL/dylib/so。32 位 Windows 使用 stdcall；Windows FFI 需要 WPS 12.1.0.23542+。其它平台仍需 WPS 发行版提供 JS 宏、UserForm 和 FFI/原生文件能力。

`tests/platform-services.test.mjs` 验证原生接口优先、没有 Crypto 的纯 JS UUID、文件错误和短读写；`tests/native-platform.test.mjs` 在 CI 三平台调用实际系统文件库，验证所有字节、中文路径、截断、打包宏修改 PPTX 与备份。Koffi 仅为测试依赖，不进入离线安装包或 PPTM。自动检查不等于已经在三平台 WPS 中完成窗体、鼠标事件及 WPS FFI 实测。

接口依据：[WPS FileSystem](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/FileSystem/obj)、[WPS JS 宏 FFI](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-macro-editor-development/run-macro/jside-ffi)。
