/* Standalone WPS JS macro. All module-level statements are declarations. */
function OpenTableStyleEditor() {
    try {
        TableStylesSlideEditor();
        return true;
    } catch (error) {
        TableStylesSlideError(error);
        return false;
    }
}

function TableStylesVersion() { return __TABLE_STYLES_VERSION__; }

function TableStylesSlideEditor() {
    if (!TableStylesSlideEditor.cached) {
        if (typeof Promise !== "function") throw new Error("当前 WPS JS 宏不支持 Promise。");
        var controls = TableStylesOpenControls(1);
        if (["打开 PPTX", "正在打开…"].indexOf(String(controls.OpenFile.Caption)) < 0) {
            throw new Error("未载入幻灯片控件。请使用完整 PPTM，并退出开发工具的设计模式。");
        }
        if (typeof globalThis.setImmediate !== "function") {
            globalThis.setImmediate = function (callback) { Promise.resolve().then(callback); };
        }
        var NativeSlideTableStyles = TableStylesNativeRuntime();
        var app = NativeSlideTableStyles.installRuntime(Application);
        var editor = new NativeSlideTableStyles.SlideEditor(app, TableStylesSlideControls(1), function () {
            var surface = new NativeSlideTableStyles.SlideColorSurface(app, editor.presentation, TableStylesPickerControls, function (error) {
                TableStylesSlideError(error);
                if (typeof MsgBox === "function") MsgBox(error.message || String(error), 0, "色盘未能显示");
            });
            return new NativeSlideTableStyles.NativeColorPicker(surface);
        });
        // Initialize can synchronously emit Change events from native controls.
        TableStylesSlideEditor.cached = editor;
        try {
            editor.initialize();
        }
        catch (error) { TableStylesSlideEditor.cached = null; throw error; }
    }
    return TableStylesSlideEditor.cached;
}

function TableStylesSlideError(error) {
    var message = error.message || String(error);
    OpenTableStyleEditor.lastError = message;
    try {
        if (TableStylesSlideEditor.cached) TableStylesSlideEditor.cached.status(message);
        else TsErrorText.Caption = message;
    }
    catch (ignored) {
        if (typeof MsgBox === "function") MsgBox(message, 0, "表格样式编辑器");
    }
}

function TableStylesSlideAction(action, page) {
    try {
        var editor = TableStylesSlideEditor();
        if (page && !editor.acceptEvent(page)) return;
        return editor.dispatch(action);
    }
    catch (error) { TableStylesSlideError(error); }
}

async function TableStylesOpenFile(page) {
    if (TableStylesOpenFile.pending) return;
    var cached = TableStylesSlideEditor.cached;
    if (cached && (cached.busy || !cached.acceptEvent(page || 1))) return;
    TableStylesOpenFile.pending = true;
    var controls;
    try {
        // Show the native chooser before loading the ZIP/XML runtime or filling
        // dropdowns. The first click should not wait for editor initialization.
        var dialog = Application.FileDialog(3);
        dialog.Title = "打开需要编辑表格样式的 PPTX";
        dialog.AllowMultiSelect = false;
        dialog.Filters.Clear(); dialog.Filters.Add("PowerPoint 演示文稿", "*.pptx");
        if (Number(dialog.Show()) !== -1) return;
        var path = String(dialog.SelectedItems.Item(1));
        controls = TableStylesOpenControls(page || 1);
        controls.OpenFile.Caption = "正在打开…"; controls.OpenFile.Enabled = false;
        controls.FirstPlaceholder.Caption = "正在初始化编辑器…";
        await Promise.resolve();
        var start = Date.now(), editor = TableStylesSlideEditor();
        var initialized = Date.now();
        controls.FirstPlaceholder.Caption = "正在读取文稿…";
        await editor.openSelected(path);
        TableStylesOpenFile.lastTimings = { initializeMs: initialized - start, loadMs: Date.now() - initialized };
    } catch (error) { TableStylesSlideError(error); }
    finally {
        if (controls) { controls.OpenFile.Caption = "打开 PPTX"; controls.OpenFile.Enabled = true; }
        if (TableStylesSlideEditor.cached) TableStylesSlideEditor.cached.enabled();
        TableStylesOpenFile.pending = false;
    }
}

function TableStylesColorAction(action, value, page) {
    try { return TableStylesSlideEditor().pickerAction(action, value, page); }
    catch (error) { TableStylesSlideError(error); }
}

function DiagnoseTableStyleEditor() {
    var report = "幻灯片宏版 " + TableStylesVersion();
    try {
        var controls = TableStylesSlideControls();
        report += "\n幻灯片控件：" + Object.keys(controls).length;
        report += "\n打开按钮：" + controls.OpenFile.Caption;
        var editor = TableStylesSlideEditor();
        report += "\n初始化：完成";
        report += "\n文件接口：" + (editor.host.app.FileSystem.backend || "WPS FileSystem");
        var platform = editor.host.app.TableStylePlatform;
        if (platform) report += "\n系统：" + platform.platform + "\nUUID：" + platform.uuidBackend;
        if (TableStylesOpenFile.lastTimings) report += "\n打开耗时：" + JSON.stringify(TableStylesOpenFile.lastTimings);
        if (editor.picker && editor.picker.form.lastPointer) report += "\n最后取色事件：" + JSON.stringify(editor.picker.form.lastPointer);
    } catch (error) { report += "\n" + (error.message || String(error)); }
    if (OpenTableStyleEditor.lastError) report += "\n" + OpenTableStyleEditor.lastError;
    if (typeof MsgBox === "function") MsgBox(report, 0, "表格样式编辑器诊断");
    return report;
}
