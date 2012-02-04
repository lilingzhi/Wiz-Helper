var objApp = WizExplorerApp
var objDatabase = objApp.Database;
var objWindow = objApp.Window;
//
var pluginPath = objApp.GetPluginPathByScriptFileName("km.js"); //获得插件的路径
objWindow.ViewHtml(pluginPath + "starstat.htm", true);
