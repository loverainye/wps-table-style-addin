/* WPS JS 宏入口。界面使用文件内的原生 UserForm，无需网页窗体。 */
// Keep the visible entry first: F5 with the initial cursor runs this function.
function OpenTableStyleEditor() {
    try {
        if (typeof TableStyleForm === "undefined") {
            throw new Error("缺少原生窗体资源。请打开完整的宏版 PPTM 文件，不能只导入 JS 代码。");
        }
        if (!(Number(TableStyleForm.Width) > 0) || !(Number(TableStyleForm.Height) > 0) ||
                String(TableStyleForm.OpenFile.Caption) !== "打开 PPTX") {
            throw new Error("WPS 未正确载入编辑器窗体及控件（窗体尺寸：" +
                TableStyleForm.Width + " × " + TableStyleForm.Height + "）。需要修正原生窗体资源，不能继续编辑文件。");
        }
        // Do not depend on the native Initialize event finding a function in
        // another module. Populate controls explicitly before showing them.
        TableStylesFormEditor();
        TableStyleForm.Show();
        if (TableStyleForm.Visible === false) {
            throw new Error("WPS 已执行窗体 Show()，但窗体仍报告为不可见。请运行 DiagnoseTableStyleEditor 查看窗体状态。");
        }
        return true;
    } catch (error) {
        var message = "表格样式编辑器启动失败：\n" + (error.message || String(error));
        OpenTableStyleEditor.lastError = message;
        if (typeof MsgBox === "function") {
            MsgBox(message, 0, "表格样式编辑器");
            return false;
        }
        throw new Error(message);
    }
}

function TableStylesVersion() {
    return __TABLE_STYLES_VERSION__;
}

function TableStylesRuntimeProbe() {
    var fs;
    try { fs = Application.FileSystem; } catch (ignored) {}
    return JSON.stringify({
        version: TableStylesVersion(),
        UserForm: typeof TableStyleForm,
        Promise: typeof Promise,
        FileSystem: !!fs,
        readFileString: fs ? typeof fs.readFileString : "undefined",
        readAsBinaryString: fs ? typeof fs.readAsBinaryString : "undefined",
        writeAsBinaryString: fs ? typeof fs.writeAsBinaryString : "undefined",
        FileDialog: typeof Application.FileDialog
    });
}

// This probe is intentionally visible when run, unlike a return-only macro.
function DiagnoseTableStyleEditor() {
    var report = TableStylesRuntimeProbe();
    try {
        report += "\n窗体类型：" + typeof TableStyleForm;
        report += "\nShow：" + typeof TableStyleForm.Show;
        report += "\nVisible：" + String(TableStyleForm.Visible);
        report += "\n尺寸：" + TableStyleForm.Width + " × " + TableStyleForm.Height;
        report += "\n打开按钮：" + typeof TableStyleForm.OpenFile;
        report += "\n编辑器已初始化：" + !!TableStylesFormEditor.cached;
        if (TableStylesFormEditor.cached) {
            var files = TableStylesFormEditor.cached.host.app.FileSystem;
            report += "\n文件接口：" + (files.backend || "WPS FileSystem");
        }
    } catch (error) {
        report += "\n窗体读取失败：" + (error.message || String(error));
    }
    if (OpenTableStyleEditor.lastError) report += "\n" + OpenTableStyleEditor.lastError;
    if (typeof MsgBox === "function") MsgBox(report, 0, "表格样式编辑器诊断");
    else throw new Error(report);
    return report;
}
