// JSA only permits declarations at module scope. Initialize the bundled
// runtime, XML parser, scheduler and controller lazily inside this function.
function TableStylesFormEditor() {
    if (!TableStylesFormEditor.cached) {
        if (typeof Promise !== "function") throw new Error("当前 WPS JS 宏不支持 Promise。");
        var root = globalThis;
        if (typeof root.setImmediate !== "function") {
            root.setImmediate = function (callback) { Promise.resolve().then(callback); };
        }
        __TABLE_STYLES_NATIVE_BUNDLE__
        var geometry = __TABLE_STYLES_NATIVE_GEOMETRY__;
        var hostApplication = NativeTableStyles.installRuntime(Application);
        var editor = new NativeTableStyles.NativeEditor(hostApplication, TableStyleForm, geometry);
        // Native Change events can run while initialize() clears/fills lists.
        // Publish the controller first so those events see updating=true
        // instead of recursively constructing another editor.
        TableStylesFormEditor.cached = editor;
        try { editor.initialize(); }
        catch (error) {
            TableStylesFormEditor.cached = null;
            throw error;
        }
    }
    return TableStylesFormEditor.cached;
}

function TableStyleForm_Initialize() {
    try { TableStylesFormEditor(); }
    catch (error) { TableStyleForm.ErrorText.Caption = error.message || String(error); }
}
