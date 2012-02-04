/*
Wiz显示插件html窗口的时候（如html对话框或者下拉框），html窗口的external对象，就是WizExplorerApp对象。
通过这个对象，可以获得当前正在运行的Wiz的一些内部对象，并调用Wiz的响应功能。
*/
//var objApp = window.external; //WizExplorerApp
var abb;
var objApp = window.external;
var objDB = objApp.Database;
var objWindow = objApp.Window;
var objDoc = objWindow.CurrentDocument; //获得当前正在浏览的Wiz文档(WizDocument)
var objHtmlDocument = objWindow.CurrentDocumentHtmlDocument;    //获得当前正在浏览的html网页的document对象(IHTMLDocument2)
//
var pluginPath = objApp.GetPluginPathByScriptFileName("km.js"); //获得插件的路径
var languangeFileName = pluginPath + "plugin.ini";  //语言文件
//
//
objApp.LocalizeHtmlDocument(languangeFileName, document);



function DateString(pDate){
    var s = pDate;
    var _str="";
    _str += s.getFullYear() + "-" + s.getMonth() + "-" + s.getDay() + " ";
    _str += s.getHours() + ":" + s.getMinutes() + ":" + s.getSeconds();
   return( _str);
}





// 把当前输入在对话框的关键字保存起来
function textKeywords_onchange() {
    objDoc.Keywords = textKeywords.value;   //保存关键字
}

function textAuthor_onchange() {
    objDoc.Author = textAuthor.value;       //保存作者
}
//
/*
通过关键字搜索相关文档，显示在文档窗口中。注：目前没有将关键字进行分隔，这样会有缺陷。
*/
function SearchByKeywords() {
    var keywords = textKeywords.value;
    if (keywords == null || keywords == "")
        return false;
    //
    try {
        keywords = keywords.replace(/\'/g, "''");
        var sql = "document_keywords like '%" + keywords + "%'";
        var documents = objDB.DocumentsFromSQL(sql);
        if (documents != null) {
            objWindow.DocumentsCtrl.SetDocuments(documents);
        }
    }
    catch (err) {
    }
}
/*
通过作者搜索相关文档，显示在文档窗口中。注：目前没有将作者进行分隔，这样会有缺陷。
    
*/
function SearchByAuthor() {
    var author = textAuthor.value;
    if (author == null || author == "")
        return false;
    //
    try {
        author = author.replace(/\'/g, "''"); // 转义
        var sql = "document_author like '%" + author + "%'";
        var documents = objDB.DocumentsFromSQL(sql);
        if (documents != null) {
            objWindow.DocumentsCtrl.SetDocuments(documents);
        }
    }
    catch (err) {
    }
}

//
/*
星标功能。我们将星标数据保存到文档的参数里面，如下面的getDocRate和setDocRate方法。
文档参数可以有任意数量，可以和服务器进行同步。
*/
//
function getDocRate() {
    var rateval = objDoc.ParamValue("Rate");
    if (rateval == null || rateval == "")
        return 0;
    //
    return parseInt(rateval);
}
function setDocRate(rateval) {
    objDoc.ParamValue("Rate") = "" + rateval;
}
//
var imgRateArray = [rate1, rate2, rate3, rate4, rate5];
var imgRateSrc1 = 'start1.gif';
var imgRateSrc2 = 'start2.gif';
//
function onRateImageClick() {
    setDocRate(window.event.srcElement._num + 1);
}
//
function onRateImageMouseOver() {
    var elem = window.event.srcElement;
    //
    for (var j = 0; j < imgRateArray.length; j++) {
        if (j <= elem._num) {
            imgRateArray[j].src = imgRateSrc2;
        } else {
            imgRateArray[j].src = imgRateSrc1;
        }
    }
}
//
for (var i = 0; i < imgRateArray.length; i++) {
    imgRateArray[i]._num = i;
    imgRateArray[i].onclick = onRateImageClick;
    imgRateArray[i].onmouseover = onRateImageMouseOver;
    imgRateArray[i].onmouseout = resetRate;
}
//
function resetRate() {
    var imgnum = getDocRate();
    for (n = 0; n < imgRateArray.length; n++) {
        imgRateArray[n].src = imgRateSrc1;
    }
    for (n = 0; n < imgnum; n++) {
        imgRateArray[n].src = imgRateSrc2;
    }
}
//
function removeRate() {
    objDoc.ParamValue("Rate") = "";
    resetRate();
}
//


function getDocGUID() {
    return objDoc.GUID;
}


function getDocTitle() {
    return objDoc.Title;
}

function getDocAuthor() {
    return objDoc.Author
}

function getDocKeyword() {
    return objDoc.Keywords;
}

function getDocLocation() {
    return objDoc.Location;
}

function getDocTag() {
    var _str = objDoc.TagsText;
    if (_str == null || _str =="") {
        _str = "未设置"
    }
    return _str;
}

function getDocOwner() {
    return objDoc.Owner;
}

function getDocCDate() {
  return objDoc.DateCreated;
}

function getDocMDate() {
    return objDoc.DateModified;
}

function getDocURL() {
    return objDoc.URL;
}

function getDocReadCount() {
    return objDoc.ReadCount-1;
}


function colorRefToHtmlColor(colorref)
{
    var _bgr= colorref.toString(16);
    var _bgr_len = _bgr.length;
    if (6 - _bgr_len <3) {
        for (var i = 1; i <= 6-_bgr_len; i++ ) {
            _bgr = "0" + _bgr;
        }
    }
    var r = _bgr.substr(4,2);
    var g = _bgr.substr(2,2);
    var b = _bgr.substr(0,2);
    return r+g+b;
}


function getDocStyle() {
    var _style;
    var _str = "";

    if ( objDoc.Style != null) {
        _style = objDoc.Style;
//      dd.innerText = colorRefToHtmlColor(_style.BackColor);
        _str = _str + "<td><span id=DocStyle><FONT style='background-color:#" + colorRefToHtmlColor(_style.BackColor) + "' ";
        _str = _str + " color=#" + colorRefToHtmlColor(_style.TextColor);
        if (_style.TextBold==1) {
            _str = _str + " ><STRONG>" + _style.Name + "</STRONG>";
        } else
            _str = _str + " > " + _style.Name;
        _str = _str + "</span></td>";
    } else {
        _str = "<td>未设置</td>"
    }

    return _str;
}

function AddInfoPage() {

    ivalue_document_guid.innerHTML = getDocGUID();
    ivalue_document_title.innerHTML = getDocTitle();

    textAuthor.value = getDocAuthor();
    textKeywords.value = getDocKeyword();

    ivalue_document_location.innerHTML = getDocLocation();
    ivalue_document_tag.innerHTML = getDocTag();
    ivalue_document_owner.innerHTML = getDocOwner();
    ivalue_dt_created.innerHTML = getDocCDate();
    ivalue_dt_modified.innerHTML = getDocMDate();
    ivalue_document_url.innerHTML = getDocURL();
    ivalue_document_style.innerHTML = getDocStyle();
    ivalue_document_readcount.innerHTML = getDocReadCount();

}


function getKeyValue(objDB, pKey) {
    var tmpStr="OFF";
    var _param, _key;

    _param = "wizhelp_parm";
    _key = pKey;
    if (_key != "" || _key != null ) {
        if (objDB.Meta(_param, _key) == "0")
            tmpStr = "ON";
        return tmpStr;
    }
}

function setKeyValue(objDB, pKey) {
    var _param, _key;
    var _value;
    _param = "WIZHELP_PARM";
    _key = pKey;
    if (_key != "" || _key != null ) {
        _value = document.getElementById(_key);
        if ( _value.value == "ON" ) {
            objDB.Meta(_param, _key) = "1";
            _value.value = "OFF";
        } else {
            objDB.Meta(_param, _key) = "0";
            _value.value = "ON";
        }
    }
}

function add_keyword_flag_button() {
    TD_KEYWORD_FLAG.innerHTML="维基语法高亮&nbsp;&nbsp;<input id=\"KEYWORD_FLAG\" type=\"button\" onclick=\"setKeyValue(objDB,\'KEYWORD_FLAG\')\" value=\"" + getKeyValue(objDB, "KEYWORD_FLAG") + "\" />";
}


function add_smarttag_flag_button() {
    TD_SMARTTAG_FLAG.innerHTML="为知助手开关&nbsp;&nbsp;<input id=\"SMARTTAG_FLAG\" type=\"button\" onclick=\"setKeyValue(objDB,\'SMARTTAG_FLAG\')\" value=\"" + getKeyValue(objDB, "SMARTTAG_FLAG") + "\" />";
}

function update_version() {
//update for V1.0.5
    if (objDB.Meta("keyword_HL","keyword_flag").value != null || 
        objDB.Meta("keyword_HL","keyword_flag").value != "" ) {
        objDB.Meta("WIZHELP_PARM","KEYWORD_FLAG").value = objDB.Meta("keyword_HL","keyword_flag").value;
//                objDB.Meta("keyword_HL","keyword_flag").delete;
    }
}


////////////////////////////////////////////////////////////////////////////////
function initPage() {
    update_version();
    resetRate();
    AddInfoPage();
    add_keyword_flag_button();
    add_smarttag_flag_button();
}


function Main() {
    initPage();
}


////////////////////////////////////////////////////////////////////////////////
//var objApp = window.external;
//var objApp = WizExplorerApp;
//var objDB = objApp.Database;
Main();

