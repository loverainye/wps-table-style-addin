/* Additional entry in the combined PPTM; no slide editor is initialized here. */
function OpenTableStyleUserForm() {
    try {
        if (typeof TableStyleForm === "undefined") throw new Error("缺少文件内的 TableStyleForm 窗体资源，请打开完整 PPTM。");
        if (!(Number(TableStyleForm.Width) > 0) || !(Number(TableStyleForm.Height) > 0) ||
                String(TableStyleForm.OpenFile.Caption) !== "打开 PPTX") {
            throw new Error("WPS 未正确载入 UserForm 窗体或控件。");
        }
        TableStylesFormEditor();
        TableStyleForm.Show();
        if (TableStyleForm.Visible === false) throw new Error("WPS 已执行 Show()，但窗体仍不可见。");
        return true;
    } catch (error) {
        TableStylesUserFormError(error);
        if (typeof MsgBox === "function") MsgBox(error.message || String(error), 0, "UserForm 启动失败");
        return false;
    }
}

function TableStylesFormEditor() {
    if (!TableStylesFormEditor.cached) {
        if (typeof Promise !== "function") throw new Error("当前 WPS JS 宏不支持 Promise。");
        if (typeof globalThis.setImmediate !== "function") {
            globalThis.setImmediate = function (callback) { Promise.resolve().then(callback); };
        }
        var runtime = TableStylesNativeRuntime();
        var geometry = __TABLE_STYLES_USERFORM_GEOMETRY__;
        var app = runtime.installRuntime(Application);
        var editor = new runtime.UserFormEditor(app, TableStyleForm, geometry, function () {
            if (typeof TableStyleColorForm === "undefined" || !(Number(TableStyleColorForm.Width) > 0)) {
                throw new Error("缺少文件内的 TableStyleColorForm 色盘窗体，请打开完整 PPTM。");
            }
            var surface = new runtime.UserFormColorSurface(TableStyleColorForm, __TABLE_STYLES_USERFORM_PICKER_GEOMETRY__, function () {
                editor.enabled();
            });
            return new runtime.NativeColorPicker(surface);
        });
        // Publish before list population can synchronously emit Change events.
        TableStylesFormEditor.cached = editor;
        try { editor.initialize(); }
        catch (error) { TableStylesFormEditor.cached = null; throw error; }
    }
    return TableStylesFormEditor.cached;
}

function TableStylesUserFormAction(action) {
    try { return TableStylesFormEditor().dispatch(action); }
    catch (error) { TableStylesUserFormError(error); }
}

function TableStylesUserFormColorAction(action, value) {
    try {
        var editor = TableStylesFormEditor.cached;
        if (!editor || !editor.picker || !editor.picker.form.visible || editor.picker.form.syncing) return;
        return editor.pickerAction(action, value);
    } catch (error) { TableStylesUserFormError(error); }
}

function TableStyleForm_Initialize() {
    try { TableStylesFormEditor(); }
    catch (error) { TableStylesUserFormError(error); }
}

function TsOpenUserForm_Click() { return OpenTableStyleUserForm(); }

function TableStylesUserFormError(error) {
    var message = error.message || String(error);
    OpenTableStyleUserForm.lastError = message;
    try { TableStyleForm.ErrorText.Caption = message; } catch (ignored) {}
}
