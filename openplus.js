	var objApp = window.external; 
	var objWindow = objApp.Window;
var objDoc = objWindow.currentdocument;
var objUI = new ActiveXObject("WizKMControls.WizCommonUI");

        var iniFile = objApp.DataStore + "wizconfig.ini";
        //var objUI = objApp.CommonUI;


function OpenIEBrowser() {

    var IEBrowser = ""; 


}

function OpenExtBrowser() {
            var ExtBrowser = ""; 
        ExtBrowser = objUI.GetValueFromIni(iniFile, "EXTBROWSER", "ExtBrowser1");


        if (ExtBrowser != null  && ExtBrowser != "") {
            objUI.RunExe(ExtBrowser, objDoc.URL, false);
        } else {
            objUI.SetValueToIni(iniFile, "EXTBROWSER", "ExtBrowser1", ExtBrowser);

        }
}


    function CloseDialog(ret) 
	{
        if (ret == 1) 
		{
            StartCommit();
        }
        else if (ret == 2) 
		{
            StartSettings();
        }
        else if (ret == 3) 
		{
            StartView();
        }		
		
        objApp.Window.CloseHtmlDialog(document, ret);
    }
