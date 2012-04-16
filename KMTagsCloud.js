var objApp = WizExplorerApp
var objWindow = objApp.Window;
//
var pluginPath = objApp.GetPluginPathByScriptFileName("KMHelper.js"); //获得插件的路径
objWindow.ViewHtml(pluginPath + "KMTagsCloud.htm", true);
