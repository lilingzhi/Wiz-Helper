var objApp = WizExplorerApp
var objWindow = objApp.Window;
//
var pluginPath = objApp.GetPluginPathByScriptFileName("km.js"); //获得插件的路径
objWindow.ViewHtml(pluginPath + "tagscloud.htm", true);
