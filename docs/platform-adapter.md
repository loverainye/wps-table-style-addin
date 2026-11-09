# 文件与 UUID 适配层

`macro/platform-services.js` 向编辑器提供 `files`、`tempPath()`、`uuid()`，并通过 `macroApplication()` 包装 WPS 宿主。不会向 WPS 的 Application 代理写入属性。

原生能力优先，各项能力独立选择，不会因为 UUID 缺少 Crypto 就放弃已有的 WPS 文件接口：

| 能力 | 优先接口 | 缺少时的实现 |
| --- | --- | --- |
| 二进制文件读写 | `Application.FileSystem`、宏全局 `FileSystem` | 当前系统的 FFI 文件接口 |
| 临时目录 | `Application.Env.GetTempPath()`、全局 `Env`、`FileSystem.tmpdir()` | Linux `/tmp`；macOS `TMPDIR`；Windows `GetTempPathW` |
| UUID v4 | JS `crypto.randomUUID()`，其次 `crypto.getRandomValues()` | Linux 系统 UUID；macOS `getentropy`；Windows `BCryptGenRandom` |

`Platform.OS()` 用于判断系统；缺少该接口且缺少原生文件能力时，依次探测可加载的系统库。不会在已识别的 Windows/macOS 上尝试 Linux 文件或 UUID 路径。新系统实现只需遵守适配层契约，编辑器和 ZIP/XML 逻辑无需改动。

| 系统 | 文件实现 | UUID 兜底 |
| --- | --- | --- |
| Linux | `macro/linux-files.js` 加载 libc，复用 POSIX 分块读写 | `/proc/sys/kernel/random/uuid` |
| macOS | `macro/platform/macos.js` 加载 `/usr/lib/libSystem.B.dylib`，使用 Darwin 的文件标志 | libSystem `getentropy` |
| Windows | `macro/platform/windows.js` 调用 kernel32 的 Unicode 文件接口，支持中文、盘符和 UNC 路径 | bcrypt `BCryptGenRandom` |

`src/uuid.js` 统一输出带花括号的大写 UUID v4。宏运行时注入平台提供者；JS 加载项直接使用浏览器 Crypto。不使用 `Math.random()` 生成样式 ID，也不需要为 UUID 启动外部程序。

FFI 为 WPS 宏自带接口，未随模板附带 DLL/dylib/so。32 位 Windows 使用 stdcall；Windows FFI 需要 WPS 12.1.0.23542+。其它平台仍需 WPS 发行版提供 JS 宏、UserForm 和 FFI/原生文件能力。

`tests/platform-services.test.mjs` 验证能力选择、原生接口优先、错误和短读写；`tests/native-platform.test.mjs` 在 CI 三平台调用实际系统库，验证所有字节、中文路径、截断、UUID、打包宏修改 PPTX 与备份。Koffi 仅为测试依赖，不进入离线安装包或 PPTM。自动检查不等于已经在三平台 WPS 中完成窗体、鼠标事件及 WPS FFI 实测。

接口依据：[WPS FileSystem](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/FileSystem/obj)、[WPS JS 宏 FFI](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-macro-editor-development/run-macro/jside-ffi)。
